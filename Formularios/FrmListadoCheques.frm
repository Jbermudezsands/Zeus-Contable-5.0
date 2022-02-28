VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FrmListadoCheques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Cheques"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   9195
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Encabezados"
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   2415
      Begin VB.CheckBox Check2 
         Caption         =   "Imprimir Cheque y Comprobante"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir Comprobante"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Consecutivo Cheque"
      Height          =   1455
      Left            =   2520
      TabIndex        =   2
      Top             =   4800
      Width           =   3255
      Begin VB.CheckBox ChkRetencion 
         Caption         =   "Imprimir Contancia Retencion"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox LblConsecutivo 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmListadoCheques.frx":0000
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc AdoCheques 
      Height          =   375
      Left            =   960
      Top             =   7320
      Width           =   4935
      _ExtentX        =   8705
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
      Caption         =   "AdoCheques"
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
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   4800
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc AdoImprime 
      Height          =   375
      Left            =   960
      Top             =   7800
      Width           =   4935
      _ExtentX        =   8705
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
      Caption         =   "AdoImprime"
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
      Left            =   960
      Top             =   8280
      Width           =   4935
      _ExtentX        =   8705
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
   Begin TrueOleDBGrid80.TDBGrid TDBGridNominas 
      Bindings        =   "FrmListadoCheques.frx":0070
      Height          =   4575
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8070
      _LayoutType     =   4
      _RowHeight      =   19
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   1
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Fecha"
      Columns(0).DataField=   "FechaTransaccion"
      Columns(0).NumberFormat=   "Short Date"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Movimiento No."
      Columns(1).DataField=   "NumeroMovimiento"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Beneficiario"
      Columns(2).DataField=   "Beneficiario"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Monto"
      Columns(3).DataField=   "Credito"
      Columns(3).DataWidth=   50
      Columns(3).NumberFormat=   "Currency"
      Columns(3).EditMask=   "##,##.##"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   80
      Columns(4)._MaxComboItems=   5
      Columns(4).ValueItems(0)._DefaultItem=   0
      Columns(4).ValueItems(0).Value=   "0"
      Columns(4).ValueItems(0).Value.vt=   8
      Columns(4).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(4).ValueItems(0).DisplayValue(0)=   "bHQAAGoIAABCTWoIAAAAAAAANgAAACgAAAAcAAAAGQAAAAEAGAAAAAAANAgAAAAAAAAAAAAAAAAA"
      Columns(4).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(8)=   "//////////////////////////////////////////////////////////////////+EhoSEhoT/"
      Columns(4).ValueItems(0).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(10)=   "//////////////////////8AAP8AAIQAAISEhoT///////////////////8AAP+EhoT/////////"
      Columns(4).ValueItems(0).DisplayValue(11)=   "//////////////////////////////////////////////////////////8AAP8AAIQAAIQAAISE"
      Columns(4).ValueItems(0).DisplayValue(12)=   "hoT///////////8AAP8AAIQAAISEhoT/////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(13)=   "//////////////////8AAP8AAIQAAIQAAIQAAISEhoT///8AAP8AAIQAAIQAAIQAAISEhoT/////"
      Columns(4).ValueItems(0).DisplayValue(14)=   "//////////////////////////////////////////////////////////8AAP8AAIQAAIQAAIQA"
      Columns(4).ValueItems(0).DisplayValue(15)=   "AISEhoQAAIQAAIQAAIQAAIQAAISEhoT/////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(16)=   "//////////////////////8AAP8AAIQAAIQAAIQAAIQAAIQAAIQAAIQAAISEhoT/////////////"
      Columns(4).ValueItems(0).DisplayValue(17)=   "//////////////////////////////////////////////////////////////8AAP8AAIQAAIQA"
      Columns(4).ValueItems(0).DisplayValue(18)=   "AIQAAIQAAIQAAISEhoT/////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(19)=   "//////////////////////////8AAIQAAIQAAIQAAIQAAISEhoT/////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(20)=   "//////////////////////////////////////////////////////////////8AAP8AAIQAAIQA"
      Columns(4).ValueItems(0).DisplayValue(21)=   "AIQAAISEhoT/////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(22)=   "//////////////////8AAP8AAIQAAIQAAIQAAIQAAISEhoT/////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(23)=   "//////////////////////////////////////////////////8AAP8AAIQAAIQAAISEhoQAAIQA"
      Columns(4).ValueItems(0).DisplayValue(24)=   "AIQAAISEhoT/////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(25)=   "//////8AAP8AAIQAAIQAAISEhoT///8AAP8AAIQAAIQAAISEhoT/////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(26)=   "//////////////////////////////////////////8AAP8AAIQAAISEhoT///////////8AAP8A"
      Columns(4).ValueItems(0).DisplayValue(27)=   "AIQAAIQAAISEhoT/////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(28)=   "//////8AAP8AAIT///////////////////8AAP8AAIQAAIQAAIT/////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(29)=   "//////////////////////////////////////////////////////////////////////////8A"
      Columns(4).ValueItems(0).DisplayValue(30)=   "AP8AAIQAAP//////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(31)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(32)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(33)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(34)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(35)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(36)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(37)=   "//////////////////////////////////////////////////////////////////////8="
      Columns(4).ValueItems(0).DisplayValue.vt=   9
      Columns(4).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(4).ValueItems(1)._DefaultItem=   0
      Columns(4).ValueItems(1).Value=   "-1"
      Columns(4).ValueItems(1).Value.vt=   8
      Columns(4).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(4).ValueItems(1).DisplayValue(0)=   "bHQAABYIAABCTRYIAAAAAAAANgAAACgAAAAcAAAAGAAAAAEAGAAAAAAA4AcAAAAAAAAAAAAAAAAA"
      Columns(4).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(8)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(10)=   "//////////////////////////////////////+EAACEAAD/////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(11)=   "//////////////////////////////////////////////////////////////////////+EAAAA"
      Columns(4).ValueItems(1).DisplayValue(12)=   "hgAAhgCEAAD/////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(13)=   "//////////////////////////+EAAAAhgAAhgAAhgAAhgCEAAD/////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(14)=   "//////////////////////////////////////////////////////////+EAAAAhgAAhgAAhgAA"
      Columns(4).ValueItems(1).DisplayValue(15)=   "hgAAhgAAhgCEAAD/////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(16)=   "//////////////+EAAAAhgAAhgAAhgAA/wAAhgAAhgAAhgAAhgCEAAD/////////////////////"
      Columns(4).ValueItems(1).DisplayValue(17)=   "//////////////////////////////////////////////////8AhgAAhgAAhgAA/wD///8A/wAA"
      Columns(4).ValueItems(1).DisplayValue(18)=   "hgAAhgAAhgCEAAD/////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(19)=   "//////////8A/wAAhgAA/wD///////////8A/wAAhgAAhgAAhgCEAAD/////////////////////"
      Columns(4).ValueItems(1).DisplayValue(20)=   "//////////////////////////////////////////////////8A/wD///////////////////8A"
      Columns(4).ValueItems(1).DisplayValue(21)=   "/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(22)=   "//////////////////////////////////////8A/wAAhgAAhgAAhgCEAAD/////////////////"
      Columns(4).ValueItems(1).DisplayValue(23)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(24)=   "//8A/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(25)=   "//////////////////////////////////////////8A/wAAhgAAhgAAhgCEAAD/////////////"
      Columns(4).ValueItems(1).DisplayValue(26)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(27)=   "//////8A/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(28)=   "//////////////////////////////////////////////8A/wAAhgAAhgCEAAD/////////////"
      Columns(4).ValueItems(1).DisplayValue(29)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(30)=   "//////////8A/wAAhgAAhgD/////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(31)=   "//////////////////////////////////////////////////8A/wD/////////////////////"
      Columns(4).ValueItems(1).DisplayValue(32)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(33)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(34)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(35)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(36)=   "//////////////////////////////////8="
      Columns(4).ValueItems(1).DisplayValue.vt=   9
      Columns(4).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(4).ValueItems.Count=   2
      Columns(4).Caption=   "Imprimir"
      Columns(4).DataField=   "ImprimeCheque"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).Caption=   "Impresion de Cheques"
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2117"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2037"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).Button=1"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2566"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=5292"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=5212"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=8196"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=2117"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2037"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=8194"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=1931"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=1852"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=1"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
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
      _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=62,.parent=13,.alignment=2"
      _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
      _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
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
End
Attribute VB_Name = "FrmListadoCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private ew As cls_NumEnglishWord
Private sw As cls_NumSpanishWord

Private Sub CmdImprimir_Click()

Dim Fechas1 As String, Fechas2 As String
Dim CodigoCuenta As String, Concepto As String
Dim x, y, H, V, Page As Integer, Dia As String, mes As String, Año As String
Dim i, J As Integer, Fechass As Date
Dim TotalDebito, TotalCredito, Totalpag As Double
Dim SubTotal, Total, IGV As Double, cadena As String
Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double, X5 As Double, Y5 As Double, X6 As Double, Y6 As Double, X7 As Double, Y7 As Double, X8 As Double, Y8 As Double, X9 As Double, Y9 As Double, X10 As Double, Y10 As Double, X11 As Double, Y11 As Double, X12 As Double, Y12 As Double, X13 As Double, Y13 As Double
Dim UltimaLinea As Double, DiferenciaY As Double, NLineas As Double
Dim Caracter As Double, ContadorLinea As Double, CadenaDescripcion As String, CaracteresLineas As Double
Dim Meses As Double, ConsecutivoCheque As Double
Dim Letras As String, Memo As String, Beneficiario As String, TipoMoneda As String, NumeroTransaccion As String, Ciudad As String
Dim CuentasContancia As String, NoConstancia As String

CodigoCuenta = FrmCheque.DBCodigo.Text
ConsecutivoCheque = Me.LblConsecutivo.Text

Sql = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo, Transacciones.Beneficiario, Transacciones.FechaVence, IndiceTransaccion.FechaTransaccion AS Expr1, IndiceTransaccion.NumeroMovimiento AS Expr2, IndiceTransaccion.ImprimeCheque,IndiceTransaccion.TipoMoneda FROM Periodos INNER JOIN  Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
      "WHERE (Transacciones.ChequeNo = '#######') AND (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (IndiceTransaccion.ImprimeCheque = 1) AND (Transacciones.NombreCuenta <> '**********CANCELADO*************') ORDER BY Transacciones.FechaTransaccion,Transacciones.NumeroMovimiento"
Me.AdoImprime.RecordSource = Sql
Me.AdoImprime.Refresh

Do While Not Me.AdoImprime.Recordset.EOF


            Beneficiario = Me.AdoImprime.Recordset("Beneficiario")
            Memo = Me.AdoImprime.Recordset("DescripcionMovimiento")
            TipoMoneda = Me.AdoImprime.Recordset("TipoMoneda")

            
            Page = 1
            
            Printer.FontSize = 6
            Printer.ScaleMode = 6
            
            
            
            TotalDebito = 0
            TotalCredito = 0
            
            
            If Not IsNumeric(Me.LblConsecutivo.Text) Then
             MsgBox "El Numero del Cheque debe Ser Numerico", vbCritical, "Sistema contable"
             Exit Sub
            End If
            
            
            '///////imprimo el reporte/////
             Debito = 0
             Credito = 0
             TotalDebito = 0
             TotalCredito = 0
                  NumFecha1 = Me.AdoImprime.Recordset("FechaTransaccion")
                  Fechas1 = Format(Me.AdoImprime.Recordset("FechaTransaccion"), "yyyymmdd")
                  NMovimiento = Val(Me.AdoImprime.Recordset("NumeroMovimiento"))
                  NumeroMovimiento = Val(Me.AdoImprime.Recordset("NumeroMovimiento"))
                  FrmCheque.DtaConsulta.RecordSource = "SELECT FechaTransaccion, CodCuentas, NTransaccion, NumeroMovimiento, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, TCambio, Debito, Credito From Transacciones WHERE (FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (NumeroMovimiento = " & NMovimiento & ")"
                  FrmCheque.DtaConsulta.Refresh
                Do While Not FrmCheque.DtaConsulta.Recordset.EOF

                   MontoCheque = Me.AdoImprime.Recordset("Credito")

                   Debito = FrmCheque.DtaConsulta.Recordset("Credito")
                   Credito = FrmCheque.DtaConsulta.Recordset("Credito")
                   TotalDebito = TotalDebito + Debito
                   TotalCredito = TotalCredito + Credito
                   FrmCheque.DtaConsulta.Recordset.MoveNext
                 Loop
                  
                    CodigoCuenta = FrmCheque.DBCodigo.Text
                    FrmCheque.DtaConsulta.RecordSource = "SELECT CodCuentas, ConsecutivoCheque From NConsecutivos WHERE (CodCuentas = '" & CodigoCuenta & "')"
                    FrmCheque.DtaConsulta.Refresh
                    If Not FrmCheque.DtaConsulta.Recordset.EOF Then
                        FrmCheque.DtaConsulta.Recordset("ConsecutivoCheque") = ConsecutivoCheque
                        FrmCheque.DtaConsulta.Recordset.Update
                    End If
              
              
                    FrmCheque.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento,Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito,Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas,Transacciones.NumeroMovimiento , Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas1 & "', 102))AND (Transacciones.NumeroMovimiento = " & NMovimiento & ")ORDER BY Transacciones.NTransaccion"
                    FrmCheque.DtaConsulta.Refresh
                    Do While Not FrmCheque.DtaConsulta.Recordset.EOF
                    
                      FrmCheque.DtaConsulta.Recordset("ChequeNo") = ConsecutivoCheque
                      FrmCheque.DtaConsulta.Recordset.Update
                      FrmCheque.DtaConsulta.Recordset.MoveNext
                    Loop
                    
                    
           
                    
            If TipoMoneda = "Dólares" Then
             Letras = sw.ConvertCurrencyToSpanish(MontoCheque, "Dólares")
            ElseIf TipoMoneda = "Córdobas" Then
             Letras = sw.ConvertCurrencyToSpanish(MontoCheque, "Córdobas")
             
            End If
            
                   
            Monto = MontoCheque
              
        If Me.Check1.Value = 1 Then
             If Dir(RutaLogo) <> "" Then
             ArepCheque.Logo.Picture = LoadPicture(RutaLogo)
             End If
'             ArepCheque.DtaCheque.ConnectionString = ConexionReporte
'             ArepCheque.LblDescripcionMonto.Caption = Letras
'             ArepCheque.LblMemo.Caption = Memo
'             ArepCheque.LblMonto.Caption = Format(MontoCheque, "##,##0.00")
'             ArepCheque.LblNombre.Caption = Beneficiario
'             ArepCheque.LblChequeNo.Caption = ConsecutivoCheque
             
             
             ArepCheque2.DtaCheque.ConnectionString = ConexionReporte
             ArepCheque.DtaCheque.ConnectionString = ConexionReporte
             
             If FrmCheque.CmbMoneda.Text = "Córdobas" Then
               If FrmCheque.ChkCheque.Value = 1 Then
               
                 TasaCambio = BuscaTasaCambio(FrmCheque.TxtFecha.Value)
                 Monto = FrmCheque.TxtMonto.Text
                 Monto = Monto / TasaCambio
                 ArepCheque.LblDescripcionMonto.Caption = sw.ConvertCurrencyToSpanish(Monto, "Dólares")
                 ArepCheque.LblMonto.Caption = Format(Monto, "##,##0.00")
                 
                 ArepCheque2.LblDescripcionMonto.Caption = sw.ConvertCurrencyToSpanish(Monto, "Dólares")
                 ArepCheque2.LblMonto.Caption = Format(Monto, "##,##0.00")
               Else
                ArepCheque.LblDescripcionMonto.Caption = FrmCheque.TxtLetras.Text
                ArepCheque.LblMonto.Caption = Format(FrmCheque.TxtMonto.Text, "##,##0.00")
                
                ArepCheque2.LblDescripcionMonto.Caption = FrmCheque.TxtLetras.Text
                ArepCheque2.LblMonto.Caption = Format(FrmCheque.TxtMonto.Text, "##,##0.00")
               End If
             Else
                ArepCheque.LblDescripcionMonto.Caption = FrmCheque.TxtLetras.Text
                ArepCheque.LblMonto.Caption = Format(FrmCheque.TxtMonto.Text, "##,##0.00")
                
                ArepCheque2.LblDescripcionMonto.Caption = FrmCheque.TxtLetras.Text
                ArepCheque2.LblMonto.Caption = Format(FrmCheque.TxtMonto.Text, "##,##0.00")
             End If
             ArepCheque.LblMemo.Caption = FrmCheque.TxtMemo.Text
             ArepCheque2.LblMemo.Caption = FrmCheque.TxtMemo.Text
             ArepCheque2.Memo = FrmCheque.TxtMemo.Text
             
             ArepCheque.LblNombre.Caption = FrmCheque.TxtNombre.Text
             ArepCheque.LblChequeNo.Caption = Me.LblConsecutivo.Text
             
             
             ArepCheque2.LblNombre.Caption = FrmCheque.TxtNombre.Text
             ArepCheque2.LblChequeNo.Caption = Me.LblConsecutivo.Text
             ArepCheque2.ChequeNo = Me.LblConsecutivo.Text
             ArepCheque2.Moneda = TipoMoneda
            
            FechaCheque = Fechas1
            NumeroMovimientos = NumeroTransaccion
            
                If TipoMoneda = "Córdobas" Then
                    
                    ArepCheque.DtaCheque.Source = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NumeroMovimiento, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.Debito, Transacciones.Credito, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTasas,  " & _
                                                  "CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito / Tasas.MontoCordobas ELSE Transacciones.Debito END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito / Tasas.MontoCordobas ELSE Transacciones.Credito END AS CreditoD, Transacciones.NPeriodo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
                                                  "WHERE (Transacciones.FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ") ORDER BY Transacciones.NTransaccion"
                    ArepCheque2.DtaCheque.Source = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NumeroMovimiento, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.Debito, Transacciones.Credito, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTasas,  " & _
                                                  "CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito / Tasas.MontoCordobas ELSE Transacciones.Debito END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito / Tasas.MontoCordobas ELSE Transacciones.Credito END AS CreditoD, Transacciones.NPeriodo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
                                                  "WHERE (Transacciones.FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (Transacciones.NumeroMovimiento = " & NMovimiento & ") ORDER BY Transacciones.NTransaccion"
                Else
                    ArepCheque.DtaCheque.Source = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NumeroMovimiento, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.Debito, Transacciones.Credito, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTasas,  " & _
                                                  "CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito * Tasas.MontoCordobas ELSE Transacciones.Debito END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito / Tasas.MontoCordobas ELSE Transacciones.Credito END AS CreditoD, Transacciones.NPeriodo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
                                                  "WHERE (Transacciones.FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ") ORDER BY Transacciones.NTransaccion"
                    ArepCheque2.DtaCheque.Source = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NumeroMovimiento, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito*Tasas.MontoCordobas  ELSE Transacciones.Debito END AS Debito,  CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito*Tasas.MontoCordobas  ELSE Transacciones.Credito END AS Credito, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTasas,  " & _
                                                  "CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito  ELSE Transacciones.Debito * Tasas.MontoCordobas END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito  ELSE Transacciones.Credito * Tasas.MontoCordobas END AS CreditoD, Transacciones.NPeriodo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
                                                  "WHERE (Transacciones.FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (Transacciones.NumeroMovimiento = " & NMovimiento & ") ORDER BY Transacciones.NTransaccion"
                
                End If
             ArepCheque2.Show 1
             
  ElseIf Me.Check2.Value = 1 Then

'---------------------------------------------------------------------------------------------------------------------------------------
'-------------------------------------IMPRIMO EL COMPROBANTE PRIMERO------------------------------------
'-------------------------------------------------------------------------------------------------------
         ArepCheque2.DtaCheque.ConnectionString = ConexionReporte
         ArepCheque.DtaCheque.ConnectionString = ConexionReporte
         
         If TipoMoneda = "Córdobas" Then
           If FrmCheque.ChkCheque.Value = 1 Then
           
             TasaCambio = BuscaTasaCambio(Me.AdoImprime.Recordset("FechaTransaccion"))
             Monto = MontoCheque
             Monto = Monto / TasaCambio
             ArepCheque.LblDescripcionMonto.Caption = sw.ConvertCurrencyToSpanish(Monto, "Dólares")
             ArepCheque.LblMonto.Caption = Format(Monto, "##,##0.00")
             
             ArepCheque2.LblDescripcionMonto.Caption = sw.ConvertCurrencyToSpanish(Monto, "Dólares")
             ArepCheque2.LblMonto.Caption = Format(Monto, "##,##0.00")
           Else
           
            Monto = MontoCheque
            ArepCheque.LblDescripcionMonto.Caption = Letras
            ArepCheque.LblMonto.Caption = Format(Monto, "##,##0.00")
            
            ArepCheque2.LblDescripcionMonto.Caption = Letras
            ArepCheque2.LblMonto.Caption = Format(Monto, "##,##0.00")
           End If
         Else
            ArepCheque.LblDescripcionMonto.Caption = Letras
            ArepCheque.LblMonto.Caption = Format(Monto, "##,##0.00")
            
            ArepCheque2.LblDescripcionMonto.Caption = Letras
            ArepCheque2.LblMonto.Caption = Format(Monto, "##,##0.00")
         End If
         
         
         ArepCheque.LblMemo.Caption = Memo
         ArepCheque2.LblMemo.Caption = Memo
         ArepCheque2.Memo = Memo
         
         ArepCheque.LblNombre.Caption = Beneficiario
         ArepCheque.LblChequeNo.Caption = Me.LblConsecutivo.Text
         
         ArepCheque2.LblNombre.Caption = Beneficiario
         ArepCheque2.LblChequeNo.Caption = Me.LblConsecutivo.Text
         ArepCheque2.ChequeNo = Me.LblConsecutivo.Text
         ArepCheque2.Moneda = TipoMoneda
        
        FechaCheque = Fechas1
        NumeroMovimientos = NMovimiento
        
            If TipoMoneda = "Córdobas" Then
                
                ArepCheque.DtaCheque.Source = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NumeroMovimiento, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.Debito, Transacciones.Credito, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTasas,  " & _
                                              "CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito / Tasas.MontoCordobas ELSE Transacciones.Debito END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito / Tasas.MontoCordobas ELSE Transacciones.Credito END AS CreditoD, Transacciones.NPeriodo, Cuentas.DescripcionGrupo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
                                              "WHERE (Transacciones.FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (Transacciones.NumeroMovimiento = " & NMovimiento & ") ORDER BY Transacciones.NTransaccion"
                ArepCheque2.DtaCheque.Source = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NumeroMovimiento, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.Debito, Transacciones.Credito, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTasas,  " & _
                                              "CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito / Tasas.MontoCordobas ELSE Transacciones.Debito END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito / Tasas.MontoCordobas ELSE Transacciones.Credito END AS CreditoD, Transacciones.NPeriodo, Cuentas.DescripcionGrupo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
                                              "WHERE (Transacciones.FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (Transacciones.NumeroMovimiento = " & NMovimiento & ") ORDER BY Transacciones.NTransaccion"
            Else
                ArepCheque.DtaCheque.Source = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NumeroMovimiento, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.Debito, Transacciones.Credito, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTasas,  " & _
                                              "CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito * Tasas.MontoCordobas ELSE Transacciones.Debito END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito / Tasas.MontoCordobas ELSE Transacciones.Credito END AS CreditoD, Transacciones.NPeriodo, Cuentas.DescripcionGrupo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
                                              "WHERE (Transacciones.FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (Transacciones.NumeroMovimiento = " & NMovimiento & ") ORDER BY Transacciones.NTransaccion"
                ArepCheque2.DtaCheque.Source = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NumeroMovimiento, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito*Tasas.MontoCordobas  ELSE Transacciones.Debito END AS Debito,  CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito*Tasas.MontoCordobas  ELSE Transacciones.Credito END AS Credito, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTasas,  " & _
                                              "CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito  ELSE Transacciones.Debito * Tasas.MontoCordobas END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito  ELSE Transacciones.Credito * Tasas.MontoCordobas END AS CreditoD, Transacciones.NPeriodo, Cuentas.DescripcionGrupo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
                                              "WHERE (Transacciones.FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (Transacciones.NumeroMovimiento = " & NMovimiento & ") ORDER BY Transacciones.NTransaccion"
            
            End If
         ArepCheque2.Show 1
 
 
        MsgBox "Coloque El Comprobante en la impresora", vbInformation, "Zeus Contable"
        
 
        FrmCheque.AdoCordenadas.RecordSource = "SELECT CodCuenta, X1, Y1, X2, Y2, X3, Y3, X4, Y4, X5, Y5, X6, Y6, X7, Y7, X8, Y8, X9, Y9, X10, Y10, X11, Y11, X12, Y12, X13, Y13,X14, Y14,X15, Y15,X16, Y16,X17, Y17, X18, Y18, X19, Y19,X20, Y20,X21, Y21, X22, Y22, NLineas,CaracteresLineas, CaracteresConcepto, Ciudad From CordenadasCheque WHERE  (CodCuenta = '" & CodigoCuenta & "')"
        FrmCheque.AdoCordenadas.Refresh
        If FrmCheque.AdoCordenadas.Recordset.EOF Then
         MsgBox "No Existen Coordenadas, para la Cuenta", vbCritical, "Sistema Contable"
         Exit Sub
        End If


        X1 = FrmCheque.AdoCordenadas.Recordset("X1")
        Y1 = FrmCheque.AdoCordenadas.Recordset("Y1")
        X2 = FrmCheque.AdoCordenadas.Recordset("X2")
        Y2 = FrmCheque.AdoCordenadas.Recordset("Y2")
        X3 = FrmCheque.AdoCordenadas.Recordset("X3")
        Y3 = FrmCheque.AdoCordenadas.Recordset("Y3")
        X4 = FrmCheque.AdoCordenadas.Recordset("X4")
        Y4 = FrmCheque.AdoCordenadas.Recordset("Y4")
        X5 = FrmCheque.AdoCordenadas.Recordset("X5")
        Y5 = FrmCheque.AdoCordenadas.Recordset("Y5")
        X6 = FrmCheque.AdoCordenadas.Recordset("X6")
        Y6 = FrmCheque.AdoCordenadas.Recordset("Y6")
        X7 = FrmCheque.AdoCordenadas.Recordset("X7")
        Y7 = FrmCheque.AdoCordenadas.Recordset("Y7")
        X8 = FrmCheque.AdoCordenadas.Recordset("X8")
        Y8 = FrmCheque.AdoCordenadas.Recordset("Y8")
        X9 = FrmCheque.AdoCordenadas.Recordset("X9")
        Y9 = FrmCheque.AdoCordenadas.Recordset("Y9")
        X10 = FrmCheque.AdoCordenadas.Recordset("X10")
        Y10 = FrmCheque.AdoCordenadas.Recordset("Y10")
        X11 = FrmCheque.AdoCordenadas.Recordset("X11")
        Y11 = FrmCheque.AdoCordenadas.Recordset("Y11")
        X12 = FrmCheque.AdoCordenadas.Recordset("X12")
        Y12 = FrmCheque.AdoCordenadas.Recordset("Y12")
        X13 = FrmCheque.AdoCordenadas.Recordset("X13")
        Y13 = FrmCheque.AdoCordenadas.Recordset("Y13")
        X14 = FrmCheque.AdoCordenadas.Recordset("X14")
        Y14 = FrmCheque.AdoCordenadas.Recordset("Y14")
        X15 = FrmCheque.AdoCordenadas.Recordset("X15")
        Y15 = FrmCheque.AdoCordenadas.Recordset("Y15")
        X16 = FrmCheque.AdoCordenadas.Recordset("X16")
        Y16 = FrmCheque.AdoCordenadas.Recordset("Y16")
        X17 = FrmCheque.AdoCordenadas.Recordset("X17")
        Y17 = FrmCheque.AdoCordenadas.Recordset("Y17")
        X18 = FrmCheque.AdoCordenadas.Recordset("X18")
        Y18 = FrmCheque.AdoCordenadas.Recordset("Y18")
        X19 = FrmCheque.AdoCordenadas.Recordset("X19")
        Y19 = FrmCheque.AdoCordenadas.Recordset("Y19")
        X20 = FrmCheque.AdoCordenadas.Recordset("X20")
        Y20 = FrmCheque.AdoCordenadas.Recordset("Y20")
        X21 = FrmCheque.AdoCordenadas.Recordset("X21")
        Y21 = FrmCheque.AdoCordenadas.Recordset("Y21")
        X22 = FrmCheque.AdoCordenadas.Recordset("X22")
        Y22 = FrmCheque.AdoCordenadas.Recordset("Y22")
        NLineas = Val(FrmCheque.AdoCordenadas.Recordset("NLineas"))
        CaracteresLineas = Val(FrmCheque.AdoCordenadas.Recordset("CaracteresLineas"))
        CaracteresConcepto = Val(FrmCheque.AdoCordenadas.Recordset("CaracteresConcepto"))
        Ciudad = FrmCheque.AdoCordenadas.Recordset("Ciudad")
        
       If TipoMoneda = "Córdobas" Then
            If FrmCheque.ChkCheque.Value = 1 Then
              TasaCambio = BuscaTasaCambio(Me.AdoImprime.Recordset("FechaTransaccion"))
              Monto = MontoCheque
              Monto = Monto / TasaCambio
              FrmCheque.TxtLetras.Text = Letras
'              ArepCheque.LblMonto.Caption = Format(Monto, "##,##0.00")
            Else
              Monto = MontoCheque
            End If
         Else
           Monto = MontoCheque
      End If
      
      Concepto = Memo
      
      
        FrmCheque.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento,Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito,Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas,Transacciones.NumeroMovimiento , Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas1 & "', 102))AND (Transacciones.NumeroMovimiento = " & NMovimiento & ")ORDER BY Transacciones.NTransaccion"
        FrmCheque.DtaConsulta.Refresh
        Printer.FontSize = 8
        'Inicio el Ciclo de Impresion
        i = 1
        
        TasaCambio = BuscaTasaCambio(FrmCheque.TxtFecha.Value)
'        Do While Not FrmCheque.DtaConsulta.Recordset.EOF
        
                   If i = 1 Then
                        ContadorLinea = i
                          
                                  '//////////////////////////////////////////////////////////////////////////////////////
                                  '//////////////////////////IMPRIMO LOS ENCABEZADOS ////////////////////////////////////
                                  '//////////////////////////////////////////////////////////////////////////////////
                           
                                   If X5 <> 0 Or Y5 <> 0 Then
                                     Caracter = 1
                                     LineaConcepto = 1
                                     cadena = Concepto
                                     If Len(cadena) > CaracteresConcepto Then
                                          Do While Len(cadena) >= CaracteresConcepto
                                                 If Caracter = 1 Then
                '                                    Printer.CurrentX = Val(X5) '5
                '                                    Printer.CurrentY = Val(Y5) + (5 * i) '120
                '                                    Printer.FontName = "Times New Roman"
                '                                    Printer.FontSize = 11
                '                                    Printer.FontBold = True
                '                                    Printer.Print Concepto
                                                    
                                                           
                                                                 cadena = Mid(Concepto, 1, CaracteresConcepto)
                                                                 Printer.CurrentX = Val(X5) '25
                                                                 Printer.CurrentY = Val(Y5) + (5 * LineaConcepto)
                                                                 Printer.FontName = "Times New Roman"
                                                                 Printer.FontSize = 11
                                                                 Printer.FontBold = True
                                                                 Printer.Print cadena
                                                                 Caracter = Caracter + CaracteresConcepto
                                                                 
                                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA SIGUIENTE/////////////////
                                                                 
                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                 If Len(cadena) < CaracteresConcepto Then
                                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                                     LineaConcepto = LineaConcepto + 1
                                                                     Printer.CurrentX = Val(X5) '25
                                                                     Printer.CurrentY = Val(Y5) + (5 * LineaConcepto)
                                                                     Printer.FontName = "Times New Roman"
                                                                     Printer.FontSize = 11
                                                                     Printer.FontBold = True
                                                                     Printer.Print cadena
                                                                     
                                                                     Caracter = Caracter + CaracteresConcepto
                                                                 End If
                                                                 
                                                 Else
                                                                 
                                                                 LineaConcepto = LineaConcepto + 1
                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                 Printer.CurrentX = Val(X5) '25
                                                                 Printer.CurrentY = Val(Y5) + (5 * LineaConcepto)
                                                                 Printer.FontName = "Times New Roman"
                                                                 Printer.FontSize = 11
                                                                 Printer.FontBold = True
                                                                 Printer.Print cadena
                                                                 
                                                                 Caracter = Caracter + CaracteresConcepto
                                                                 
                                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA/////////////////
                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                 If Len(cadena) < CaracteresConcepto Then
                                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                                     LineaConcepto = LineaConcepto + 1
                                                                     Printer.CurrentX = Val(X5) '25
                                                                     Printer.CurrentY = Val(Y5) + (5 * LineaConcepto)
                                                                     Printer.FontName = "Times New Roman"
                                                                     Printer.FontSize = 11
                                                                     Printer.FontBold = True
                                                                     Printer.Print cadena
                                                                     
                                                                     Caracter = Caracter + CaracteresConcepto
                                                                 End If
                                                    
                                                 End If
                                          Loop
                                          
                                     Else
                                                    Printer.CurrentX = Val(X5) '5
                                                    Printer.CurrentY = Val(Y5) + (5 * i) '120
                                                    Printer.FontName = "Times New Roman"
                                                    Printer.FontSize = 11
                                                    Printer.FontBold = True
                                                    Printer.Print Concepto
                                     End If
                                   End If
                                   
                                   
                                  If X18 <> 0 Or Y18 <> 0 Then
                                     Caracter = 1
                                     LineaConcepto = 1
                                     cadena = Memo
                                     If Len(cadena) > CaracteresConcepto Then
                                          Do While Len(cadena) >= CaracteresConcepto
                                                 If Caracter = 1 Then
                '                                    Printer.CurrentX = Val(X5) '5
                '                                    Printer.CurrentY = Val(Y5) + (5 * i) '120
                '                                    Printer.FontName = "Times New Roman"
                '                                    Printer.FontSize = 11
                '                                    Printer.FontBold = True
                '                                    Printer.Print Concepto
                                                    
                                                           
                                                                 cadena = Mid(Concepto, 1, CaracteresConcepto)
                                                                 Printer.CurrentX = Val(X18) '25
                                                                 Printer.CurrentY = Val(Y18) + (5 * LineaConcepto)
                                                                 Printer.FontName = "Times New Roman"
                                                                 Printer.FontSize = 11
                                                                 Printer.FontBold = True
                                                                 Printer.Print cadena
                                                                 Caracter = Caracter + CaracteresConcepto
                                                                 
                                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA SIGUIENTE/////////////////
                                                                 
                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                 If Len(cadena) < CaracteresConcepto Then
                                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                                     LineaConcepto = LineaConcepto + 1
                                                                     Printer.CurrentX = Val(X18) '25
                                                                     Printer.CurrentY = Val(Y18) + (5 * LineaConcepto)
                                                                     Printer.FontName = "Times New Roman"
                                                                     Printer.FontSize = 11
                                                                     Printer.FontBold = True
                                                                     Printer.Print cadena
                                                                     
                                                                     Caracter = Caracter + CaracteresConcepto
                                                                 End If
                                                                 
                                                 Else
                                                                 
                                                                 LineaConcepto = LineaConcepto + 1
                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                 Printer.CurrentX = Val(X18) '25
                                                                 Printer.CurrentY = Val(Y18) + (5 * LineaConcepto)
                                                                 Printer.FontName = "Times New Roman"
                                                                 Printer.FontSize = 11
                                                                 Printer.FontBold = True
                                                                 Printer.Print cadena
                                                                 
                                                                 Caracter = Caracter + CaracteresConcepto
                                                                 
                                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA/////////////////
                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                 If Len(cadena) < CaracteresConcepto Then
                                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                                     LineaConcepto = LineaConcepto + 1
                                                                     Printer.CurrentX = Val(X18) '25
                                                                     Printer.CurrentY = Val(Y18) + (5 * LineaConcepto)
                                                                     Printer.FontName = "Times New Roman"
                                                                     Printer.FontSize = 11
                                                                     Printer.FontBold = True
                                                                     Printer.Print cadena
                                                                     
                                                                     Caracter = Caracter + CaracteresConcepto
                                                                 End If
                                                    
                                                 End If
                                          Loop
                                          
                                     Else
                                                    Printer.CurrentX = Val(X18) '5
                                                    Printer.CurrentY = Val(Y18) + (5 * i) '120
                                                    Printer.FontName = "Times New Roman"
                                                    Printer.FontSize = 11
                                                    Printer.FontBold = True
                                                    Printer.Print Concepto
                                     End If
                                   End If
                                    
                                    Dia = Day(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                    mes = Month(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                    Año = Year(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                    Meses = Month(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                  
                                '    FrmCheque.DtaConsulta.Recordset.MoveLast
                                   If X9 <> 0 Or Y9 <> 0 Then
                                    Printer.CurrentX = X9
                                    Printer.CurrentY = Y9
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 11
                                    Printer.FontBold = True
                '                    Printer.Print FrmCheque.DtaConsulta.Recordset("NumeroMovimiento")
                                    Printer.Print Me.LblConsecutivo.Text
                                   End If
                                    
                                   If X1 <> 0 Or Y1 <> 0 Then
                                    Printer.CurrentX = Val(X1)
                                    Printer.CurrentY = Val(Y1) + (5 * i)
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 9
                                    Printer.FontBold = True
                                    Printer.Print Beneficiario
                                   End If
                                   
                                  If X14 <> 0 Or Y14 <> 0 Then
                                    Printer.CurrentX = Val(X14)
                                    Printer.CurrentY = Val(Y14) + (5 * i)
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 9
                                    Printer.FontBold = True
                                    Printer.Print Beneficiario
                                   End If
                                   
                                   If X4 <> 0 Or Y4 <> 0 Then
                                    Printer.CurrentX = Val(X4)
                                    Printer.CurrentY = Val(Y4) + (5 * i)
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 9
                                    Printer.FontBold = True
                                    Printer.Print Letras
                                   End If
                                   
                                  If X15 <> 0 Or Y15 <> 0 Then
                                    Printer.CurrentX = Val(X15)
                                    Printer.CurrentY = Val(Y15) + (5 * i)
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 9
                                    Printer.FontBold = True
                                    Printer.Print Letras
                                   End If
                                   
                                   If X3 <> 0 Or Y3 <> 0 Then
                                    Printer.CurrentX = Val(X3)
                                    Printer.CurrentY = Val(Y3) + (5 * i)
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 11
                                    Printer.FontBold = True
                                    Printer.Print Format(Monto, "##,##0.00")
                                   End If
                                   
                                                      
                                   If X16 <> 0 Or Y16 <> 0 Then
                                    Printer.CurrentX = Val(X16)
                                    Printer.CurrentY = Val(Y16) + (5 * i)
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 11
                                    Printer.FontBold = True
                                    Printer.Print Format(Monto, "##,##0.00")
                                   End If
                                
                                   If X2 <> 0 Or Y2 <> 0 Then
                                    Printer.CurrentX = Val(X2) '20
                                    Printer.CurrentY = Val(Y2) '288
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 11
                                    Printer.FontBold = True
                                    FechaLetra = Ciudad & "         " & Format(Day(FrmCheque.DtaConsulta.Recordset("FechaTransaccion")), "0#") & "          " & Format(Month(FrmCheque.DtaConsulta.Recordset("FechaTransaccion")), "0#") & "           " & Year(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                    Printer.Print FechaLetra
                                   End If
                                
                                    If X17 <> 0 Or Y17 <> 0 Then
                                    Printer.CurrentX = Val(X17) '20
                                    Printer.CurrentY = Val(Y17) '288
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 11
                                    Printer.FontBold = True
                                    FechaLetra = Ciudad & "          " & Format(Day(FrmCheque.DtaConsulta.Recordset("FechaTransaccion")), "0#") & "          " & Format(Month(FrmCheque.DtaConsulta.Recordset("FechaTransaccion")), "0#") & "           " & Year(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                    Printer.Print FechaLetra
                                   End If
                                
                   End If

        
        Printer.EndDoc
        
'        ConsecutivoCheque = Me.LblConsecutivo.Text
'        ConsecutivoCheque = ConsecutivoCheque + 1
'        Me.LblConsecutivo.Text = ConsecutivoCheque
'         FrmCheque.DtaConsulta.Recordset.MoveNext
'        Loop
             
             
  Else
            
           
             FrmCheque.AdoCordenadas.RecordSource = "SELECT CodCuenta, X1, Y1, X2, Y2, X3, Y3, X4, Y4, X5, Y5, X6, Y6, X7, Y7, X8, Y8, X9, Y9, X10, Y10, X11, Y11, X12, Y12, X13, Y13,X14, Y14,X15, Y15,X16, Y16,X17, Y17, X18, Y18, X19, Y19,X20, Y20,X21, Y21, X22, Y22, NLineas,CaracteresLineas, CaracteresConcepto From CordenadasCheque WHERE  (CodCuenta = '" & CodigoCuenta & "')"
             FrmCheque.AdoCordenadas.Refresh
             If FrmCheque.AdoCordenadas.Recordset.EOF Then
              MsgBox "No Existen Coordenadas, para la Cuenta", vbCritical, "Sistema Contable"
              Exit Sub
             End If
             
             
                    X1 = FrmCheque.AdoCordenadas.Recordset("X1")
                    Y1 = FrmCheque.AdoCordenadas.Recordset("Y1")
                    X2 = FrmCheque.AdoCordenadas.Recordset("X2")
                    Y2 = FrmCheque.AdoCordenadas.Recordset("Y2")
                    X3 = FrmCheque.AdoCordenadas.Recordset("X3")
                    Y3 = FrmCheque.AdoCordenadas.Recordset("Y3")
                    X4 = FrmCheque.AdoCordenadas.Recordset("X4")
                    Y4 = FrmCheque.AdoCordenadas.Recordset("Y4")
                    X5 = FrmCheque.AdoCordenadas.Recordset("X5")
                    Y5 = FrmCheque.AdoCordenadas.Recordset("Y5")
                    X6 = FrmCheque.AdoCordenadas.Recordset("X6")
                    Y6 = FrmCheque.AdoCordenadas.Recordset("Y6")
                    X7 = FrmCheque.AdoCordenadas.Recordset("X7")
                    Y7 = FrmCheque.AdoCordenadas.Recordset("Y7")
                    X8 = FrmCheque.AdoCordenadas.Recordset("X8")
                    Y8 = FrmCheque.AdoCordenadas.Recordset("Y8")
                    X9 = FrmCheque.AdoCordenadas.Recordset("X9")
                    Y9 = FrmCheque.AdoCordenadas.Recordset("Y9")
                    X10 = FrmCheque.AdoCordenadas.Recordset("X10")
                    Y10 = FrmCheque.AdoCordenadas.Recordset("Y10")
                    X11 = FrmCheque.AdoCordenadas.Recordset("X11")
                    Y11 = FrmCheque.AdoCordenadas.Recordset("Y11")
                    X12 = FrmCheque.AdoCordenadas.Recordset("X12")
                    Y12 = FrmCheque.AdoCordenadas.Recordset("Y12")
                    X13 = FrmCheque.AdoCordenadas.Recordset("X13")
                    Y13 = FrmCheque.AdoCordenadas.Recordset("Y13")
                    X14 = FrmCheque.AdoCordenadas.Recordset("X14")
                    Y14 = FrmCheque.AdoCordenadas.Recordset("Y14")
                    X15 = FrmCheque.AdoCordenadas.Recordset("X15")
                    Y15 = FrmCheque.AdoCordenadas.Recordset("Y15")
                    X16 = FrmCheque.AdoCordenadas.Recordset("X16")
                    Y16 = FrmCheque.AdoCordenadas.Recordset("Y16")
                    X17 = FrmCheque.AdoCordenadas.Recordset("X17")
                    Y17 = FrmCheque.AdoCordenadas.Recordset("Y17")
                    X18 = FrmCheque.AdoCordenadas.Recordset("X18")
                    Y18 = FrmCheque.AdoCordenadas.Recordset("Y18")
                    X19 = FrmCheque.AdoCordenadas.Recordset("X19")
                    Y19 = FrmCheque.AdoCordenadas.Recordset("Y19")
                    X20 = FrmCheque.AdoCordenadas.Recordset("X20")
                    Y20 = FrmCheque.AdoCordenadas.Recordset("Y20")
                    X21 = FrmCheque.AdoCordenadas.Recordset("X21")
                    Y21 = FrmCheque.AdoCordenadas.Recordset("Y21")
                    X22 = FrmCheque.AdoCordenadas.Recordset("X22")
                    Y22 = FrmCheque.AdoCordenadas.Recordset("Y22")
                    NLineas = Val(FrmCheque.AdoCordenadas.Recordset("NLineas"))
                    CaracteresLineas = Val(FrmCheque.AdoCordenadas.Recordset("CaracteresLineas"))
                    CaracteresConcepto = Val(FrmCheque.AdoCordenadas.Recordset("CaracteresConcepto"))
             
            'Cargo la Consulta del Cheque
             FrmCheque.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento,Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito,Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas,Transacciones.NumeroMovimiento , Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas1 & "', 102))AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ")ORDER BY Transacciones.NTransaccion"
             FrmCheque.DtaConsulta.Refresh
             Printer.FontSize = 9
             'Inicio el Ciclo de Impresion
             i = 1
             Do While Not FrmCheque.DtaConsulta.Recordset.EOF
               
               '////////////////////////////////////////////////////////////////////////////////////////////
               '//////////////////////SI EL NUMERO DE LINEAS ES MAYOR CREO UNA NUEVA PAGINA////////////////
               '///////////////////////////////////////////////////////////////////////////////////////////
                
                If i > NLineas Then
                  Printer.NewPage
                  i = 1
                End If
                      
                      If i = 1 Then
                      ContadorLinea = i
                      
                              '//////////////////////////////////////////////////////////////////////////////////////
                              '//////////////////////////IMPRIMO LOS ENCABEZADOS ////////////////////////////////////
                              '//////////////////////////////////////////////////////////////////////////////////
                       
'                               If X5 <> 0 Or Y5 <> 0 Then
'                                Printer.CurrentX = Val(X5) '5
'                                Printer.CurrentY = Val(Y5) + (5 * i) '120
'                                Printer.FontName = "Times New Roman"
'                                Printer.FontSize = 11
'                                Printer.FontBold = True
'                                Printer.Print Me.AdoImprime.Recordset("DescripcionMovimiento")
'                               End If

                                       If X5 <> 0 Or Y5 <> 0 Then
                                         Caracter = 1
                                         LineaConcepto = 1
                                         cadena = Concepto
                                         If Len(cadena) > CaracteresConcepto Then
                                              Do While Len(cadena) >= CaracteresConcepto
                                                     If Caracter = 1 Then
                    '                                    Printer.CurrentX = Val(X5) '5
                    '                                    Printer.CurrentY = Val(Y5) + (5 * i) '120
                    '                                    Printer.FontName = "Times New Roman"
                    '                                    Printer.FontSize = 11
                    '                                    Printer.FontBold = True
                    '                                    Printer.Print Concepto
                                                        
                                                               
                                                                     cadena = Mid(Concepto, 1, CaracteresConcepto)
                                                                     Printer.CurrentX = Val(X5) '25
                                                                     Printer.CurrentY = Val(Y5) + (5 * LineaConcepto)
                                                                     Printer.FontName = "Times New Roman"
                                                                     Printer.FontSize = 11
                                                                     Printer.FontBold = True
                                                                     Printer.Print cadena
                                                                     Caracter = Caracter + CaracteresConcepto
                                                                     
                                                                     '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA SIGUIENTE/////////////////
                                                                     
                                                                     cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                     If Len(cadena) < CaracteresConcepto Then
                                                                      '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                                         LineaConcepto = LineaConcepto + 1
                                                                         Printer.CurrentX = Val(X5) '25
                                                                         Printer.CurrentY = Val(Y5) + (5 * LineaConcepto)
                                                                         Printer.FontName = "Times New Roman"
                                                                         Printer.FontSize = 11
                                                                         Printer.FontBold = True
                                                                         Printer.Print cadena
                                                                         
                                                                         Caracter = Caracter + CaracteresConcepto
                                                                     End If
                                                                     
                                                     Else
                                                                     
                                                                     LineaConcepto = LineaConcepto + 1
                                                                     cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                     Printer.CurrentX = Val(X5) '25
                                                                     Printer.CurrentY = Val(Y5) + (5 * LineaConcepto)
                                                                     Printer.FontName = "Times New Roman"
                                                                     Printer.FontSize = 11
                                                                     Printer.FontBold = True
                                                                     Printer.Print cadena
                                                                     
                                                                     Caracter = Caracter + CaracteresConcepto
                                                                     
                                                                     '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA/////////////////
                                                                     cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                     If Len(cadena) < CaracteresConcepto Then
                                                                      '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                                         LineaConcepto = LineaConcepto + 1
                                                                         Printer.CurrentX = Val(X5) '25
                                                                         Printer.CurrentY = Val(Y5) + (5 * LineaConcepto)
                                                                         Printer.FontName = "Times New Roman"
                                                                         Printer.FontSize = 11
                                                                         Printer.FontBold = True
                                                                         Printer.Print cadena
                                                                         
                                                                         Caracter = Caracter + CaracteresConcepto
                                                                     End If
                                                        
                                                     End If
                                              Loop
                                              
                                           Else
                                                        Printer.CurrentX = Val(X5) '5
                                                        Printer.CurrentY = Val(Y5) + (5 * i) '120
                                                        Printer.FontName = "Times New Roman"
                                                        Printer.FontSize = 11
                                                        Printer.FontBold = True
                                                        Printer.Print Concepto
                                           End If
                                         End If
                               
                               
                                      If X18 <> 0 Or Y18 <> 0 Then
                                                     Caracter = 1
                                                     LineaConcepto = 1
                                                     cadena = Concepto
                                                     If Len(cadena) > CaracteresConcepto Then
                                                          Do While Len(cadena) >= CaracteresConcepto
                                                                 If Caracter = 1 Then
                                '                                    Printer.CurrentX = Val(X5) '5
                                '                                    Printer.CurrentY = Val(Y5) + (5 * i) '120
                                '                                    Printer.FontName = "Times New Roman"
                                '                                    Printer.FontSize = 11
                                '                                    Printer.FontBold = True
                                '                                    Printer.Print Concepto
                                                                    
                                                                           
                                                                                 cadena = Mid(Concepto, 1, CaracteresConcepto)
                                                                                 Printer.CurrentX = Val(X18) '25
                                                                                 Printer.CurrentY = Val(Y18) + (5 * LineaConcepto)
                                                                                 Printer.FontName = "Times New Roman"
                                                                                 Printer.FontSize = 11
                                                                                 Printer.FontBold = True
                                                                                 Printer.Print cadena
                                                                                 Caracter = Caracter + CaracteresConcepto
                                                                                 
                                                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA SIGUIENTE/////////////////
                                                                                 
                                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                                 If Len(cadena) < CaracteresConcepto Then
                                                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                                                     LineaConcepto = LineaConcepto + 1
                                                                                     Printer.CurrentX = Val(X18) '25
                                                                                     Printer.CurrentY = Val(Y18) + (5 * LineaConcepto)
                                                                                     Printer.FontName = "Times New Roman"
                                                                                     Printer.FontSize = 11
                                                                                     Printer.FontBold = True
                                                                                     Printer.Print cadena
                                                                                     
                                                                                     Caracter = Caracter + CaracteresConcepto
                                                                                 End If
                                                                                 
                                                                 Else
                                                                                 
                                                                                 LineaConcepto = LineaConcepto + 1
                                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                                 Printer.CurrentX = Val(X18) '25
                                                                                 Printer.CurrentY = Val(Y18) + (5 * LineaConcepto)
                                                                                 Printer.FontName = "Times New Roman"
                                                                                 Printer.FontSize = 11
                                                                                 Printer.FontBold = True
                                                                                 Printer.Print cadena
                                                                                 
                                                                                 Caracter = Caracter + CaracteresConcepto
                                                                                 
                                                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA/////////////////
                                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                                 If Len(cadena) < CaracteresConcepto Then
                                                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                                                     LineaConcepto = LineaConcepto + 1
                                                                                     Printer.CurrentX = Val(X18) '25
                                                                                     Printer.CurrentY = Val(Y18) + (5 * LineaConcepto)
                                                                                     Printer.FontName = "Times New Roman"
                                                                                     Printer.FontSize = 11
                                                                                     Printer.FontBold = True
                                                                                     Printer.Print cadena
                                                                                     
                                                                                     Caracter = Caracter + CaracteresConcepto
                                                                                 End If
                                                                    
                                                                 End If
                                                          Loop
                                                          
                                                     Else
                                                                    Printer.CurrentX = Val(X18) '5
                                                                    Printer.CurrentY = Val(Y18) + (5 * i) '120
                                                                    Printer.FontName = "Times New Roman"
                                                                    Printer.FontSize = 11
                                                                    Printer.FontBold = True
                                                                    Printer.Print Concepto
                                                     End If
                                                   End If
                               
                               
                                
                                Dia = Day(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                mes = Month(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                Año = Year(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                Meses = Month(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                              
                            '    FrmCheque.DtaConsulta.Recordset.MoveLast

                               If X9 <> 0 Or Y9 <> 0 Then
                                Printer.CurrentX = X9
                                Printer.CurrentY = Y9
                                Printer.FontName = "Times New Roman"
                                Printer.FontSize = 11
                                Printer.FontBold = True
            '                    Printer.Print FrmCheque.DtaConsulta.Recordset("NumeroMovimiento")
                                Printer.Print Me.LblConsecutivo.Text
                               End If
                                
                               If X1 <> 0 Or Y1 <> 0 Then
                                Printer.CurrentX = Val(X1)
                                Printer.CurrentY = Val(Y1) + (5 * i)
                                Printer.FontName = "Times New Roman"
                                Printer.FontSize = 11
                                Printer.FontBold = True
                                Printer.Print FrmCheque.TxtNombre.Text
                               End If
                               
                              If X14 <> 0 Or Y14 <> 0 Then
                                Printer.CurrentX = Val(X14)
                                Printer.CurrentY = Val(Y14) + (5 * i)
                                Printer.FontName = "Times New Roman"
                                Printer.FontSize = 11
                                Printer.FontBold = True
                                Printer.Print FrmCheque.TxtNombre.Text
                               End If
                               
                               If X4 <> 0 Or Y4 <> 0 Then
                                Printer.CurrentX = Val(X4)
                                Printer.CurrentY = Val(Y4) + (5 * i)
                                Printer.FontName = "Times New Roman"
                                Printer.FontSize = 9
                                Printer.FontBold = True
                                Printer.Print FrmCheque.TxtLetras.Text
                               End If
                               
                              If X15 <> 0 Or Y15 <> 0 Then
                                Printer.CurrentX = Val(X15)
                                Printer.CurrentY = Val(Y15) + (5 * i)
                                Printer.FontName = "Times New Roman"
                                Printer.FontSize = 9
                                Printer.FontBold = True
                                Printer.Print FrmCheque.TxtLetras.Text
                               End If
                               
                               If X3 <> 0 Or Y3 <> 0 Then
                                Printer.CurrentX = Val(X3)
                                Printer.CurrentY = Val(Y3) + (5 * i)
                                Printer.FontName = "Times New Roman"
                                Printer.FontSize = 11
                                Printer.FontBold = True
                                Printer.Print Format(Monto, "##,##0.00")
                               End If
                               
                                                  
                               If X16 <> 0 Or Y16 <> 0 Then
                                Printer.CurrentX = Val(X16)
                                Printer.CurrentY = Val(Y16) + (5 * i)
                                Printer.FontName = "Times New Roman"
                                Printer.FontSize = 11
                                Printer.FontBold = True
                                Printer.Print Format(Monto, "##,##0.00")
                               End If
                            
                               If X2 <> 0 Or Y2 <> 0 Then
                                Printer.CurrentX = Val(X2) '20
                                Printer.CurrentY = Val(Y2) '288
                                Printer.FontName = "Times New Roman"
                                Printer.FontSize = 11
                                Printer.FontBold = True
                                FechaLetra = "Juigalpa          " & Format(Day(FrmCheque.DtaConsulta.Recordset("FechaTransaccion")), "0#") & "          " & Format(Month(FrmCheque.DtaConsulta.Recordset("FechaTransaccion")), "0#") & "           " & Year(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                Printer.Print FechaLetra
                               End If
                            
                               If X17 <> 0 Or Y17 <> 0 Then
                                Printer.CurrentX = Val(X17) '20
                                Printer.CurrentY = Val(Y17) '288
                                Printer.FontName = "Times New Roman"
                                Printer.FontSize = 11
                                Printer.FontBold = True
                                FechaLetra = "Juigalpa          " & Format(Day(FrmCheque.DtaConsulta.Recordset("FechaTransaccion")), "0#") & "          " & Format(Month(FrmCheque.DtaConsulta.Recordset("FechaTransaccion")), "0#") & "           " & Year(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                Printer.Print FechaLetra
                               End If
                
         

                            
                            
                       End If
                       
                       '//////////////////////////////////////////////////////////////////////////////////////
                       '//////////////////////////IMPRIMO LOS DETALLES ////////////////////////////////////
                       '//////////////////////////////////////////////////////////////////////////////////
                       
                       
                       If X6 <> 0 Or Y6 <> 0 Then
                        Printer.CurrentX = Val(X6) '5
                        Printer.CurrentY = Val(Y6) + (5 * i)
                        cadena = FrmCheque.DtaConsulta.Recordset("CodCuentas")
                        If Len(cadena) > 20 Then
                         cadena = Mid(cadena, 1, 20)
                        End If
                        
                        Printer.FontName = "Times New Roman"
                        Printer.FontSize = 9
                        Printer.FontBold = False
                        Printer.Print cadena
                       End If
                    
                    
                    
                    
                      If X10 <> 0 Or Y10 <> 0 Then
                        Printer.CurrentX = Val(X10) '25
                        Printer.CurrentY = Val(Y10) + (5 * i)
                        cadena = FrmCheque.DtaConsulta.Recordset("NombreCuenta")
                        If Len(cadena) > 24 Then
                         cadena = Mid(cadena, 1, 24)
                        End If
                        
                        Printer.FontName = "Times New Roman"
                        Printer.FontSize = 9
                        Printer.FontBold = False
                        Printer.Print cadena
                      End If
                    
                     
                        If X11 <> 0 Or Y11 <> 0 Then
                                 CadenaDescripcion = FrmCheque.DtaConsulta.Recordset("DescripcionMovimiento")
                                 cadena = FrmCheque.DtaConsulta.Recordset("DescripcionMovimiento")
                                 Caracter = 1
                                 ContadorLinea = i
                                 
                                 If Len(cadena) > CaracteresLineas Then
                                          Do While Len(cadena) >= CaracteresLineas
                                                   If Caracter = 1 Then
                                                             cadena = Mid(cadena, 1, CaracteresLineas)
                                                             Printer.CurrentX = Val(X11) '25
                                                             Printer.CurrentY = Val(Y11) + (5 * i)
                                                             Printer.FontName = "Times New Roman"
                                                             Printer.FontSize = 9
                                                             Printer.FontBold = False
                                                             Printer.Print cadena
                                                             Caracter = Caracter + CaracteresLineas
                                                             
                                                             '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA SIGUIENTE/////////////////
                                                             cadena = Mid(CadenaDescripcion, Caracter, CaracteresLineas)
                                                             If Len(cadena) < CaracteresLineas Then
                                                              '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                                 ContadorLinea = ContadorLinea + 1
                                                                 Printer.CurrentX = Val(X11) '25
                                                                 Printer.CurrentY = Val(Y11) + (5 * ContadorLinea)
                                                                 Printer.FontName = "Times New Roman"
                                                                 Printer.FontSize = 9
                                                                 Printer.FontBold = False
                                                                 Printer.Print cadena
                                                                 
                                                                 Caracter = Caracter + CaracteresLineas
                                                             End If
                                                     Else
                                                             ContadorLinea = ContadorLinea + 1
                                                             cadena = Mid(CadenaDescripcion, Caracter, CaracteresLineas)
                                                             Printer.CurrentX = Val(X11) '25
                                                             Printer.CurrentY = Val(Y11) + (5 * ContadorLinea)
                                                             Printer.FontName = "Times New Roman"
                                                             Printer.FontSize = 9
                                                             Printer.FontBold = False
                                                             Printer.Print cadena
                                                             
                                                             Caracter = Caracter + CaracteresLineas
                                                             
                                                             '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA/////////////////
                                                             cadena = Mid(CadenaDescripcion, Caracter, CaracteresLineas)
                                                             If Len(cadena) < CaracteresLineas Then
                                                              '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                                 ContadorLinea = ContadorLinea + 1
                                                                 Printer.CurrentX = Val(X11) '25
                                                                 Printer.CurrentY = Val(Y11) + (5 * ContadorLinea)
                                                                 Printer.FontName = "Times New Roman"
                                                                 Printer.FontSize = 9
                                                                 Printer.FontBold = False
                                                                 Printer.Print cadena
                                                                 
                                                                 Caracter = Caracter + CaracteresLineas
                                                             End If
                                                             
                                                             
                                                    End If
                             
                                                          
                                          Loop
                                                          
                                 Else
                                         Printer.CurrentX = Val(X11) '25
                                         Printer.CurrentY = Val(Y11) + (5 * i)
                                         Printer.FontName = "Times New Roman"
                                         Printer.FontSize = 9
                                         Printer.FontBold = False
                                         Printer.Print cadena
                                                       
                                 End If
                              
            
                        End If
               
                          
                    
                      
                       If X12 <> 0 Or Y12 <> 0 Then
                        Printer.CurrentX = Val(X12) '135
                         Printer.CurrentY = Val(Y12) + (5 * i)
                         Printer.FontName = "Times New Roman"
                         Printer.FontSize = 9
                         Printer.FontBold = False
                         Printer.Print Format(DebitoCordobas, "##,##0.00")
            '            Printer.Print Format(FrmCheque.DtaConsulta.Recordset("Debito"), "##,##0.00")
                       End If
                       
                          If X19 <> 0 Or Y19 <> 0 Then
                            Printer.CurrentX = Val(X19) '135
                             Printer.CurrentY = Val(Y19) + (5 * i)
                             Printer.FontName = "Times New Roman"
                             Printer.FontSize = 9
                             Printer.FontBold = False
                             Printer.Print Format(DebitoDolares, "##,##0.00")
                '            Printer.Print Format(FrmCheque.DtaConsulta.Recordset("Debito"), "##,##0.00")
                           End If
                        
                        
                           If X13 <> 0 Or Y13 <> 0 Then
                            Printer.CurrentX = Val(X13) '165
                              Printer.CurrentY = Val(Y13) + (5 * i) '165
                            Printer.FontName = "Times New Roman"
                            Printer.FontSize = 9
                            Printer.FontBold = False
                            Printer.Print Format(CreditoCordobas, "##,##0.00")
                '            Printer.Print Format(FrmCheque.DtaConsulta.Recordset("Credito"), "##,##0.00")
                           End If
                           
                           If X20 <> 0 Or Y20 <> 0 Then
                            Printer.CurrentX = Val(X20) '165
                              Printer.CurrentY = Val(Y20) + (5 * i) '165
                            Printer.FontName = "Times New Roman"
                            Printer.FontSize = 9
                            Printer.FontBold = False
                            Printer.Print Format(CreditoDolares, "##,##0.00")
                '            Printer.Print Format(FrmCheque.DtaConsulta.Recordset("Credito"), "##,##0.00")
                           End If
                       
                       
                        If i > 1 Then
                          UltimaLinea = UltimaLinea + (5 * i) + DiferenciaY - 4
                        End If
                       
                        i = ContadorLinea
                        i = i + 1
                        ContadorLinea = i
                    '
                    
                    ' 'Fin del Ciclo
            
            
                 FrmCheque.DtaConsulta.Recordset.MoveNext
             Loop
             
             

                
              If FrmCheque.CmbMoneda.Text = "Córdobas" Then
                 TotalDebitoCordobas = TotalDebito
                 TotalDebitoDolares = TotalDebitoCordobas / TasaCambio
                 
                 TotalCreditoCordobas = TotalCredito
                 TotalCreditoDolares = TotalCreditoCordobas / TasaCambio
               
              Else
                 TotalDebitoDolares = TotalDebito
                 TotalDebitoCordobas = TotalDebitoCordobas * TasaCambio
                 
                 TotalCreditoDolares = TotalCredito
                 TotalCreditoCordobas = TotalCreditoDolares * TasaCambio
                 
            
              End If
                
                 If X7 <> 0 Or Y7 <> 0 Then
                   Printer.CurrentX = Val(X7) '135
                   Printer.CurrentY = Val(Y7) '288
                   Printer.FontName = "Times New Roman"
                   Printer.FontSize = 9
                   Printer.FontBold = False
                   Printer.Print Format(TotalDebitoCordobas, "##,##0.00")
                 End If
                 
                If X21 <> 0 Or Y21 <> 0 Then
                   Printer.CurrentX = Val(X21) '135
                   Printer.CurrentY = Val(Y21) '288
                   Printer.FontName = "Times New Roman"
                   Printer.FontSize = 9
                   Printer.FontBold = False
                   Printer.Print Format(TotalDebitoDolares, "##,##0.00")
                 End If
                 
                 If X8 <> 0 Or Y8 <> 0 Then
                   Printer.CurrentX = Val(X8) '165
                   Printer.CurrentY = Val(Y8) '288
                   Printer.FontName = "Times New Roman"
                   Printer.FontSize = 9
                   Printer.FontBold = False
                   Printer.Print Format(TotalCreditoCordobas, "##,##0.00")
                 End If
                 
                If X22 <> 0 Or Y22 <> 0 Then
                   Printer.CurrentX = Val(X22) '165
                   Printer.CurrentY = Val(Y22) '288
                   Printer.FontName = "Times New Roman"
                   Printer.FontSize = 9
                   Printer.FontBold = False
                   Printer.Print Format(TotalCreditoDolares, "##,##0.00")
                 End If
            
            
             
            'termino de imprimir las facturas
            Printer.EndDoc
       End If
       
       
       
   If Me.ChkRetencion.Value = 1 Then
     '/////////////////////////////////////////////////////////////////////////////////////////////
     '///////////////////////////////////BUSCO EL CONSECUTIVO DE LA CONSTANCIA ///////////////////
     '/////////////////////////////////////////////////////////////////////////////////////////////
      CodigoCuenta = FrmCheque.DBCodigo.Text
          FrmCheque.DtaConsulta.RecordSource = "SELECT * From NConsecutivos " 'WHERE (CodCuentas = '" & CodigoCuenta & "')
          FrmCheque.DtaConsulta.Refresh
                    If Not FrmCheque.DtaConsulta.Recordset.EOF Then
                        FrmCheque.DtaConsulta.Recordset("ConstanciaRetencion") = FrmCheque.DtaConsulta.Recordset("ConstanciaRetencion") + 1
                        FrmCheque.DtaConsulta.Recordset.Update
                        NoConstancia = Format(FrmCheque.DtaConsulta.Recordset("ConstanciaRetencion"), "0000#")
                    End If
   
   
      FrmCheque.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo, Cuentas.CausaRetencion, CASE WHEN Cuentas.Cedula IS NULL THEN CASE WHEN Cuentas.RUC IS NULL THEN '00-000000-0000H' ELSE Cuentas.RUC END ELSE Cuentas.RUC END AS RUC,  Cuentas.DescRetencion  FROM  Periodos INNER JOIN  Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo INNER JOIN  Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas  " & _
                                           "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (Transacciones.NumeroMovimiento = " & NMovimiento & ") AND (Cuentas.CausaRetencion = 1) ORDER BY Transacciones.NTransaccion"
      FrmCheque.DtaConsulta.Refresh
      Do While Not FrmCheque.DtaConsulta.Recordset.EOF
      
            ArepConstanciaRetencion.LblFecha.Text = Format(Me.AdoImprime.Recordset("FechaTransaccion"), "dd/mm/yyyy")
            ArepConstanciaRetencion.LblTransaccion.Text = NMovimiento
            ArepConstanciaRetencion.LblNombre.Caption = Beneficiario
            ArepConstanciaRetencion.LblMemo.Caption = Memo
            ArepConstanciaRetencion.LblNumeroRuc.Caption = FrmCheque.DtaConsulta.Recordset("RUC")
'            ArepConstanciaRetencion.LblMonto.Caption = MontoCheque + FrmCheque.DtaConsulta.Recordset("Credito")
            ArepConstanciaRetencion.LblMontoRetencion.Caption = FrmCheque.DtaConsulta.Recordset("Credito")
            
            If TipoMoneda = "Dólares" Then
             Letras = sw.ConvertCurrencyToSpanish(FrmCheque.DtaConsulta.Recordset("Credito"), "Dólares")
            ElseIf TipoMoneda = "Córdobas" Then
             Letras = sw.ConvertCurrencyToSpanish(FrmCheque.DtaConsulta.Recordset("Credito"), "Córdobas")
             
            End If
            ArepConstanciaRetencion.LblConstanciaNo.Caption = NoConstancia
            ArepConstanciaRetencion.LblDescripcionMonto.Caption = Letras
            ArepConstanciaRetencion.LblTasaRetencion.Caption = FrmCheque.DtaConsulta.Recordset("DescRetencion")
            ArepConstanciaRetencion.Show 1
            
        FrmCheque.DtaConsulta.Recordset.MoveNext
        
        MsgBox "Impresion Correcta", vbInformation, "Zeus Contable"
      Loop
   
   End If
       
            
              
 ConsecutivoCheque = Me.LblConsecutivo.Text
 ConsecutivoCheque = ConsecutivoCheque + 1
 Me.LblConsecutivo.Text = ConsecutivoCheque
  Me.AdoImprime.Recordset.MoveNext
  Loop



'--------------------------------ACTUALIZACION DEL GRID //////////////////////////////////////////////
CodigoCuenta = FrmCheque.DBCodigo.Text

Sql = "SELECT  Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo,Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo, Transacciones.Beneficiario, Transacciones.FechaVence, IndiceTransaccion.FechaTransaccion AS Expr1, IndiceTransaccion.NumeroMovimiento AS Expr2, IndiceTransaccion.ImprimeCheque FROM  Periodos INNER JOIN  Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento  " & _
      "WHERE (Transacciones.ChequeNo = N'#######') AND (Transacciones.CodCuentas = '" & CodigoCuenta & "') ORDER BY Transacciones.FechaTransaccion,Transacciones.NumeroMovimiento"
Me.AdoCheques.RecordSource = Sql
Me.AdoCheques.Refresh

Me.TDBGridNominas.Columns(0).Button = False


   Me.LblConsecutivo.Text = ConsecutivoCheque




End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub



Private Sub Form_Initialize()
On Error GoTo TipoErrs
Dim SqlCheque As String
    Set ew = New cls_NumEnglishWord
    Set sw = New cls_NumSpanishWord
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub Form_Load()
Dim Sql As String, CodigoCuenta As String

With Me.AdoCheques
   .ConnectionString = Conexion
End With

With Me.AdoImprime
   .ConnectionString = Conexion
End With

MDIPrimero.Skin1.ApplySkin hWnd

CodigoCuenta = FrmCheque.DBCodigo.Text

Sql = "SELECT  Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo,Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo, Transacciones.Beneficiario, Transacciones.FechaVence, IndiceTransaccion.FechaTransaccion AS Expr1, IndiceTransaccion.NumeroMovimiento AS Expr2, IndiceTransaccion.ImprimeCheque FROM  Periodos INNER JOIN  Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento  " & _
      "WHERE (Transacciones.ChequeNo = N'#######') AND (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (Transacciones.NombreCuenta <> '**********CANCELADO*************') ORDER BY Transacciones.FechaTransaccion,Transacciones.NumeroMovimiento"
Me.AdoCheques.RecordSource = Sql
Me.AdoCheques.Refresh

Me.TDBGridNominas.Columns(0).Button = False


End Sub
