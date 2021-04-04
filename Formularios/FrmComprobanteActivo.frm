VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FrmComprobanteActivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprobantes Depreciacion"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFDECE&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   12855
      TabIndex        =   1
      Top             =   0
      Width           =   12855
      Begin VB.Image Image2 
         Height          =   960
         Left            =   240
         Picture         =   "FrmComprobanteActivo.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   12840
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
         Caption         =   "Comprobante Contabilizacion Activo FIJO"
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
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   5505
      End
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGridFacturacion 
      Bindings        =   "FrmComprobanteActivo.frx":C042
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8070
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
End
Attribute VB_Name = "FrmComprobanteActivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
