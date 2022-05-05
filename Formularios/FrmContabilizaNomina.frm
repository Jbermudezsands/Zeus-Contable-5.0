VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmContabilizaNomina 
   Caption         =   "Contabilizar Nominas"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12810
   ScaleHeight     =   7035
   ScaleWidth      =   12810
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc AdoCuentas 
      Height          =   375
      Left            =   480
      Top             =   8640
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFDECE&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   12975
      TabIndex        =   0
      Top             =   -120
      Width           =   12975
      Begin VB.Image Image2 
         Height          =   960
         Left            =   240
         Picture         =   "FrmContabilizaNomina.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   12960
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
         Caption         =   "Contabilizando el Sistema de Nominas"
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
         Left            =   4200
         TabIndex        =   1
         Top             =   360
         Width           =   5505
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FrmContabilizaNomina.frx":C042
      Top             =   0
   End
   Begin XtremeSuiteControls.PushButton CmdSalir 
      Height          =   375
      Left            =   11520
      TabIndex        =   2
      Top             =   6600
      Width           =   1215
      _Version        =   786432
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salir"
      UseVisualStyle  =   -1  'True
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Nominas"
      TabPicture(0)   =   "FrmContabilizaNomina.frx":26786F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TDBGridNominas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "GroupBox1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   4695
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2295
         _Version        =   786432
         _ExtentX        =   4048
         _ExtentY        =   8281
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin TrueOleDBList80.TDBCombo TDBCombo1 
            Bindings        =   "FrmContabilizaNomina.frx":26788B
            Height          =   315
            Left            =   240
            TabIndex        =   17
            Top             =   2520
            Width           =   1935
            _ExtentX        =   3413
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
            ComboStyle      =   2
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
            ListField       =   "CodTipoNomina"
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
            _PropDict       =   $"FrmContabilizaNomina.frx":2678A8
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
         Begin XtremeSuiteControls.CheckBox ChkCheques 
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   1680
            Width           =   1935
            _Version        =   786432
            _ExtentX        =   3413
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Crear Cheque x Emp"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
         Begin XtremeSuiteControls.RadioButton OptNominas 
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Nominas"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdContabilizar 
            Height          =   375
            Left            =   360
            TabIndex        =   6
            Top             =   3720
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Contabilizar"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdConsultar 
            Height          =   375
            Left            =   360
            TabIndex        =   7
            Top             =   3240
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Consultar"
            UseVisualStyle  =   -1  'True
         End
         Begin MSComCtl2.DTPicker DTPicker5 
            Height          =   345
            Left            =   840
            TabIndex        =   8
            Top             =   4200
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            Format          =   65601537
            CurrentDate     =   40301
         End
         Begin XtremeSuiteControls.RadioButton RadioButton1 
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Nominas Vacaciones"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RadioButton2 
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Nominas 13vo Mes"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RadioButton3 
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Despido y Renuncias"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Nomina:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label LblFecha 
            Caption         =   "Feha:"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   4200
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGridNominas 
         Bindings        =   "FrmContabilizaNomina.frx":267952
         Height          =   4575
         Left            =   2520
         TabIndex        =   11
         Top             =   600
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   8070
         _LayoutType     =   4
         _RowHeight      =   19
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   1
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NumNomina"
         Columns(0).DataField=   "NumNomina"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nomina"
         Columns(1).DataField=   "Nomina"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "FechaINI"
         Columns(2).DataField=   "FechaNominaINI"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "FechaFIN"
         Columns(3).DataField=   "FechaNomina"
         Columns(3).DataWidth=   50
         Columns(3).NumberFormat=   "Short Date"
         Columns(3).EditMask=   "##,##.##"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "TotalSalarioBasico"
         Columns(4).DataField=   "TotalSalarioBasico"
         Columns(4).NumberFormat=   "Standard"
         Columns(4).EditMask=   "##,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "TotalDestajo"
         Columns(5).DataField=   "TotalDestajo"
         Columns(5).NumberFormat=   "Standard"
         Columns(5).EditMask=   "##,##0.00"
         Columns(5).EditMaskUpdate=   -1  'True
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
         Columns(6).ValueItems(1).Value=   "-1"
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
         Columns(6).Caption=   "Contabilizar"
         Columns(6).DataField=   "Marca"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).Caption=   "Nominas"
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1931"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1852"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
         Splits(0)._ColumnProps(5)=   "Column(0).Button=1"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2566"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=8196"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=1931"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1852"
         Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=8196"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=2117"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2037"
         Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=8194"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=2117"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2037"
         Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=8194"
         Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(27)=   "Column(5).Width=2646"
         Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=2566"
         Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=8194"
         Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(32)=   "Column(6).Width=1931"
         Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=1852"
         Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=1"
         Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
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
         _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(67)  =   "Named:id=33:Normal"
         _StyleDefs(68)  =   ":id=33,.parent=0"
         _StyleDefs(69)  =   "Named:id=34:Heading"
         _StyleDefs(70)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   ":id=34,.wraptext=-1"
         _StyleDefs(72)  =   "Named:id=35:Footing"
         _StyleDefs(73)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(74)  =   "Named:id=36:Selected"
         _StyleDefs(75)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(76)  =   "Named:id=37:Caption"
         _StyleDefs(77)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(78)  =   "Named:id=38:HighlightRow"
         _StyleDefs(79)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(80)  =   "Named:id=39:EvenRow"
         _StyleDefs(81)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(82)  =   "Named:id=40:OddRow"
         _StyleDefs(83)  =   ":id=40,.parent=33"
         _StyleDefs(84)  =   "Named:id=41:RecordSelector"
         _StyleDefs(85)  =   ":id=41,.parent=34"
         _StyleDefs(86)  =   "Named:id=42:FilterBar"
         _StyleDefs(87)  =   ":id=42,.parent=33"
      End
   End
   Begin XtremeSuiteControls.ProgressBar osProgress1 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   6600
      Visible         =   0   'False
      Width           =   11295
      _Version        =   786432
      _ExtentX        =   19923
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin MSAdodcLib.Adodc AdoNominas 
      Height          =   450
      Left            =   480
      Top             =   7560
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   794
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
      Caption         =   "AdoNominas"
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
   Begin MSAdodcLib.Adodc AdoDatosEmpresa 
      Height          =   375
      Left            =   3720
      Top             =   8280
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
      Caption         =   "AdoDatosEmpresa"
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
   Begin MSAdodcLib.Adodc AdoTipoNominas 
      Height          =   450
      Left            =   480
      Top             =   8160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   794
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
      Caption         =   "AdoTipoNominas"
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
   Begin MSAdodcLib.Adodc AdoConsultas 
      Height          =   450
      Left            =   9360
      Top             =   9120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   794
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
      Caption         =   "AdoConsultas"
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
   Begin MSAdodcLib.Adodc AdoProcesos 
      Height          =   375
      Left            =   7080
      Top             =   8640
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
      Caption         =   "AdoProcesos"
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
   Begin MSAdodcLib.Adodc AdoDetalleNomina 
      Height          =   450
      Left            =   7440
      Top             =   7560
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   794
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
      Caption         =   "AdoDetalleNomina"
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
   Begin MSAdodcLib.Adodc AdoIcentivos 
      Height          =   375
      Left            =   7320
      Top             =   9000
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
      Caption         =   "AdoIncentivos"
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
   Begin MSAdodcLib.Adodc AdoDeducciones 
      Height          =   375
      Left            =   10320
      Top             =   8640
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
      Caption         =   "AdoDeducciones"
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
   Begin MSAdodcLib.Adodc AdoSubsidios 
      Height          =   375
      Left            =   10560
      Top             =   7680
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
      Caption         =   "AdoSubsidios"
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
   Begin MSAdodcLib.Adodc AdoConsultaNomina 
      Height          =   450
      Left            =   3720
      Top             =   8640
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   794
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
      Caption         =   "AdoConsultaNomina"
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
   Begin MSAdodcLib.Adodc AdoEmpleados 
      Height          =   450
      Left            =   3840
      Top             =   8400
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   794
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
      Caption         =   "AdoEmpleados"
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
   Begin MSAdodcLib.Adodc AdoNominaSubsidio 
      Height          =   450
      Left            =   8400
      Top             =   8400
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   794
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
      Caption         =   "AdoNominaSubsidio"
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
   Begin MSAdodcLib.Adodc AdoBuscaNomina 
      Height          =   450
      Left            =   3840
      Top             =   7560
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   794
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
      Caption         =   "AdoEmpleados"
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
End
Attribute VB_Name = "FrmContabilizaNomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ConexionNominas As String

Public Sub CmdConsultar_Click()
  Dim TipoNomina As String, CodTipoNomina As String
  
  If Me.TDBCombo1.Columns(0).Text = "" Then
   MsgBox "Necesita Seleccionar un tipo de Nomina", vbCritical, "Zeus Nominas"
   Exit Sub
  End If
  
  CodTipoNomina = Me.TDBCombo1.Columns(0).Text
 
 If Me.OptNominas.Value = True Then
   TipoNomina = "Nominas"
    Me.AdoNominas.RecordSource = "SELECT  Nomina.NumNomina, TipoNomina.Nomina, Nomina.FechaNominaINI, Nomina.FechaNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.Marca FROM  Nomina INNER JOIN TipoNomina ON Nomina.CodTipoNomina = TipoNomina.CodTipoNomina " & _
                              "WHERE (Nomina.CodTipoNomina = '" & CodTipoNomina & "') AND (Nomina.Contabilizado = 0)"
    Me.AdoNominas.Refresh
    
    If Not Me.AdoNominas.Recordset.EOF Then
        Me.CmdContabilizar.Enabled = True
        Me.LblFecha.Visible = True
        Me.DTPicker5.Value = Now
        Me.DTPicker5.Visible = True
       
    End If
 End If
 

 
 


End Sub

Private Sub CmdContabilizar_Click()

  Dim TipoFactura As String, MonedaNomina As String, Reg As Double
  Dim CodEmpleado As String, CtaSueldos As String, CtaProvAguinaldo As String, CtaProvVacaciones As String, CtaOtrosIngresos As String, CtaHorasExtra As String, CtaINSSPatronal As String, CtaINATEC As String, CtaAguinaldoPagar As String, CtaVacacionesPagar As String, CtaINSSPagar As String, CtaINATECPagar As String, CtaIRPagar As String, CtaPrestamoPagar As String, CtaNominaPagar As String, CtaInssPatronalPagar As String
  Dim MontoSueldos As Double, MontoProvAguinaldo As Double, MontoProvVacaciones As Double, MontoOtrosIngresos As Double, MontoHorasExtra As Double, MontoINSSPatronal As Double, MontoINATEC As Double, MontoAguinaldoPagar As Double, MontoVacacionesPagar As Double, MontoINSSPagar As Double, MontoINATECPagar As Double, MontoIrPatronal As Double, MontoIRPagar As Double, MontoPrestamoPagar As Double, MontoNominaPagar As Double, MontoInssPatronalPagar As Double
  Dim MontoIncentivos() As Double, MontoDeducciones() As Double, MontoSubsidios() As Double
  Dim CtaIncentivos() As String, CtaDeducciones() As String, CtaSubsidios() As String, NumNomina As Double, i As Double, Registros As Double
  Dim ExisteCodigo As Boolean, Directorio As String, NumeroPeriodo As Double, NumeroTransaccion As Double, DescripcionMovimiento As String, NumeroFactura As String, Descuento As Double
  Dim CodigoCuentaCliente As String, TotalIncentivos As Double, TotalDeducciones As Double, CtaBancos As String, CtaSubsidio As String
  Dim NombreEmpleado As String, Credito As Double, MontoSubsidio As Double, CodSubsidio As String

  DoEvents

  MonedaNomina = "Córdobas"
  CodTipoNomina = Me.TDBCombo1.Columns(0).Text
  
  
            If Me.OptNominas.Value = True Then
             TipoFactura = "Nominas"
            End If

                Select Case TipoFactura
                   Case "Nominas"
                     SqlString = "SELECT Nomina.NumNomina, TipoNomina.Nomina, Nomina.FechaNomina, Nomina.Marca, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.OtrosIngresos, DetalleNomina.Incentivos, DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones AS Salario, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.INATEC, DetalleNomina.Mes13, DetalleNomina.Vacaciones, DetalleNomina.Deducciones + DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR AS DeduccionesTotal FROM Nomina INNER JOIN TipoNomina ON Nomina.CodTipoNomina = TipoNomina.CodTipoNomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina  " & _
                                                  "WHERE (Nomina.CodTipoNomina = '" & CodTipoNomina & "') AND (Nomina.Contabilizado = 0) AND (Nomina.Marca = 1)"
                     Me.AdoProcesos.RecordSource = SqlString
                
                    Me.AdoNominas.RecordSource = "SELECT  Nomina.NumNomina, TipoNomina.Nomina, Nomina.FechaNominaINI, Nomina.FechaNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.Marca FROM  Nomina INNER JOIN TipoNomina ON Nomina.CodTipoNomina = TipoNomina.CodTipoNomina " & _
                              "WHERE (Nomina.CodTipoNomina = '" & CodTipoNomina & "') AND (Nomina.Contabilizado = 0) AND (Nomina.Marca = 1)"
                    Me.AdoNominas.Refresh
                 
                 
                 End Select
                 
                 
                 
                 
     Directorio = App.Path + "\Cuentas.txt"
     Open Directorio For Output As #1
        Print #1, "Zeus Contable"
        Print #1, "Contabilizar Nominas"
        Print #1, ""


            Me.AdoNominas.Refresh
            Do While Not Me.AdoNominas.Recordset.EOF

                         NumNomina = Me.AdoNominas.Recordset("NumNomina")
                         
                             
                             Me.AdoConsultaNomina.RecordSource = "SELECT  * From Nomina WHERE (NumNomina = " & NumNomina & ")"
                             Me.AdoConsultaNomina.Refresh
                             If Not Me.AdoConsultaNomina.Recordset.EOF Then
                               If Me.AdoConsultaNomina.Recordset("Cerrada") = False Then
                                 MsgBox "La Nomina no se ha Cerrado", vbCritical, "Zeus Contable"
                                 Close #1
                                 Exit Sub
                               End If
                             End If

                         
                             
                             Me.AdoProcesos.Refresh
                             Me.osProgress1.Visible = True
                             Me.osProgress1.Min = 0
                             Me.osProgress1.Value = 0
                             If Not Me.AdoProcesos.Recordset.EOF Then
                             Me.AdoProcesos.Recordset.MoveFirst
                             Me.osProgress1.Max = Me.AdoProcesos.Recordset.RecordCount
                             End If
                             Me.AdoProcesos.Refresh
                             
                             Reg = 1
                             Do While Not Me.AdoProcesos.Recordset.EOF
                             
                                CodEmpleado = Me.AdoProcesos.Recordset("Codempleado")
                                ExisteCodigo = True
                                
                                '//////////////////////////////////////////////////////////////////////////////////////////
                                '///////////////////////////////////BUSCO LAS CUENTAS CONTABLES ///////////////////////////
                                '/////////////////////////////////////////////////////////////////////////////////////////
                                SqlString = "SELECT  * From Historico Where (CodEmpleado = " & CodEmpleado & ")"
                                Me.AdoCuentas.RecordSource = SqlString
                                Me.AdoCuentas.Refresh
                                If Not Me.AdoCuentas.Recordset.EOF Then
                                    If Not IsNull(Me.AdoCuentas.Recordset("CuentaSueldos")) Then
                                      CtaSueldos = Me.AdoCuentas.Recordset("CuentaSueldos")
                                    End If
                                    If ValidarCuentas(CtaSueldos) = False Then Print #1, CtaSueldos; ExisteCodigo = False
                                    If Not IsNull(Me.AdoCuentas.Recordset("ProvAguinaldo")) Then
                                    CtaProvAguinaldo = Me.AdoCuentas.Recordset("ProvAguinaldo")
                                    End If
                                    If ValidarCuentas(CtaProvAguinaldo) = False Then Print #1, CtaProvAguinaldo; ExisteCodigo = False
                                    If Not IsNull(Me.AdoCuentas.Recordset("ProvVacaciones")) Then CtaProvVacaciones = Me.AdoCuentas.Recordset("ProvVacaciones")
                                    If ValidarCuentas(CtaProvVacaciones) = False Then Print #1, CtaProvVacaciones; ExisteCodigo = False
                                    If Not IsNull(Me.AdoCuentas.Recordset("CuentaOtrosIngresos")) Then CtaOtrosIngresos = Me.AdoCuentas.Recordset("CuentaOtrosIngresos")
                                    If ValidarCuentas(CtaOtrosIngresos) = False Then Print #1, CtaOtrosIngresos; ExisteCodigo = False
                                    If Not IsNull(Me.AdoCuentas.Recordset("CuentaHorasExtra")) Then CtaHorasExtra = Me.AdoCuentas.Recordset("CuentaHorasExtra")
                                    If ValidarCuentas(CtaHorasExtra) = False Then Print #1, CtaHorasExtra; ExisteCodigo = False
                                    If Not IsNull(Me.AdoCuentas.Recordset("INSSPatronal")) Then CtaINSSPatronal = Me.AdoCuentas.Recordset("INSSPatronal")
                                    If ValidarCuentas(CtaINSSPatronal) = False Then Print #1, CtaINSSPatronal; ExisteCodigo = False
                                    If Not IsNull(Me.AdoCuentas.Recordset("INATEC")) Then CtaINATEC = Me.AdoCuentas.Recordset("INATEC")
                                    If ValidarCuentas(CtaINATEC) = False Then Print #1, CtaINATEC; ExisteCodigo = False
                                    If Not IsNull(Me.AdoCuentas.Recordset("AguinaldoxPagar")) Then CtaAguinaldoPagar = Me.AdoCuentas.Recordset("AguinaldoxPagar")
                                    If ValidarCuentas(CtaAguinaldoPagar) = False Then Print #1, CtaAguinaldoPagar; ExisteCodigo = False
                                    If Not IsNull(Me.AdoCuentas.Recordset("VacacionesxPagar")) Then CtaVacacionesPagar = Me.AdoCuentas.Recordset("VacacionesxPagar")
                                    If ValidarCuentas(CtaVacacionesPagar) = False Then Print #1, CtaVacacionesPagar; ExisteCodigo = False
                                    If Not IsNull(Me.AdoCuentas.Recordset("INSSxPagar")) Then CtaINSSPagar = Me.AdoCuentas.Recordset("INSSxPagar")
                                    If ValidarCuentas(CtaINSSPagar) = False Then Print #1, CtaINSSPagar; ExisteCodigo = False
                                    If Not IsNull(Me.AdoCuentas.Recordset("INATECxPagar")) Then CtaINATECPagar = Me.AdoCuentas.Recordset("INATECxPagar")
                                    If ValidarCuentas(CtaINATECPagar) = False Then Print #1, CtaINATECPagar; ExisteCodigo = False
                                    If Not IsNull(Me.AdoCuentas.Recordset("IRxPagar")) Then CtaIRPagar = Me.AdoCuentas.Recordset("IRxPagar")
                                    If ValidarCuentas(CtaIRPagar) = False Then Print #1, CtaIRPagar; ExisteCodigo = False
                                    If Not IsNull(Me.AdoCuentas.Recordset("PrestamoxPagar")) Then CtaPrestamoPagar = Me.AdoCuentas.Recordset("PrestamoxPagar")
                                    If ValidarCuentas(CtaPrestamoPagar) = False Then Print #1, CtaPrestamoPagar; ExisteCodigo = False
                                    If Not IsNull(Me.AdoCuentas.Recordset("NominaxPagar")) Then CtaNominaPagar = Me.AdoCuentas.Recordset("NominaxPagar")
                                    If ValidarCuentas(CtaNominaPagar) = False Then Print #1, CtaNominaPagar; ExisteCodigo = False
                                    If Not IsNull(Me.AdoCuentas.Recordset("INSSPatronalPagar")) Then CtaInssPatronalPagar = Me.AdoCuentas.Recordset("INSSPatronalPagar")
                                    If ValidarCuentas(CtaInssPatronalPagar) = False Then Print #1, CtaInssPatronalPagar; ExisteCodigo = False
                                    If Not IsNull(Me.AdoCuentas.Recordset("CuentaBanco")) Then CtaBancos = Me.AdoCuentas.Recordset("CuentaBanco")
                                    If ValidarCuentas(CtaBancos) = False Then Print #1, CtaBancos; ExisteCodigo = False
'                                    CtaSubsidio = Me.AdoCuentas.Recordset("CuentaSubsidio")
                                    If ValidarCuentas(CtaSubsidio) = False Then Print #1, CtaSubsidio; ExisteCodigo = False
                                End If
                                
                                
                                '////////////////////////////BUSCO LAS CUENTAS DE DEDUCCIONES ///////////////////////////////////////
                                If Me.AdoProcesos.Recordset("DeduccionesTotal") <> 0 Then
                                    Me.AdoDeducciones.RecordSource = "SELECT  DetalleDeduccion.*, TipoDeduccion.CuentaContable, TipoDeduccion.Deduccion, Deduccion.CodEmpleado FROM  DetalleDeduccion INNER JOIN Deduccion ON DetalleDeduccion.NumDeduccion = Deduccion.NumDeduccion INNER JOIN TipoDeduccion ON Deduccion.CodTipoDeduccion = TipoDeduccion.CodTipoDeduccion  " & _
                                                      "WHERE (DetalleDeduccion.NumNomina = " & NumNomina & ") AND (Deduccion.CodEmpleado = " & CodEmpleado & ") AND (DetalleDeduccion.Valor <> 0)"
                                    Me.AdoDeducciones.Refresh
                                    i = 0
                                    Registros = Me.AdoDeducciones.Recordset.RecordCount
                                    ReDim CtaDeducciones(Registros)
                                    ReDim MontoDeducciones(Registros)
                                    If Not Me.AdoDeducciones.Recordset.EOF Then
                                    Me.AdoDeducciones.Recordset.MoveFirst
                                    End If
                                    Do While Not Me.AdoDeducciones.Recordset.EOF
                                     If Not IsNull(Me.AdoDeducciones.Recordset("CuentaContable")) Then
                                       CtaDeducciones(i) = Me.AdoDeducciones.Recordset("CuentaContable")
                                     Else
                                       MsgBox "Deducciones, No tienen Cuenta Contable" & Me.AdoDeducciones.Recordset("CodEmpleado"), vbCritical, "Zeus Contable"
                                       Exit Sub
                                     End If
                                     If ValidarCuentas(CtaDeducciones(i)) = False Then Print #1, CtaDeducciones(i); ExisteCodigo = False
                                     MontoDeducciones(i) = Me.AdoDeducciones.Recordset("Valor")
                                     i = i + 1
                                     Me.AdoDeducciones.Recordset.MoveNext
                                    Loop
                                End If

                                '////////////////////////////BUSCO LAS CUENTAS DE INCENTIVOS ///////////////////////////////////////
                                If Me.AdoProcesos.Recordset("Incentivos") <> 0 Then
                                    Me.AdoIcentivos.RecordSource = "SELECT * FROM DetalleIncentivo INNER JOIN Incentivo ON DetalleIncentivo.NumIncentivo = Incentivo.NumIncentivo INNER JOIN TipoIncentivo ON Incentivo.CodTipoIncentivo = TipoIncentivo.CodTipoIncentivo  " & _
                                                                   "Where (DetalleIncentivo.NumNomina = " & NumNomina & ") And (Incentivo.CodEmpleado = " & CodEmpleado & ") And (DetalleIncentivo.Valor <> 0)"
                                    Me.AdoIcentivos.Refresh
                                    i = 0
                                    Registros = Me.AdoIcentivos.Recordset.RecordCount
                                    ReDim CtaIncentivos(Registros)
                                    ReDim MontoIncentivos(Registros)
                                    If Not Me.AdoIcentivos.Recordset.EOF Then
                                      Me.AdoIcentivos.Recordset.MoveFirst
                                    End If
                                    Do While Not Me.AdoIcentivos.Recordset.EOF
                                    If Not IsNull(Me.AdoIcentivos.Recordset("CuentaContable")) Then
                                     CtaIncentivos(i) = Me.AdoIcentivos.Recordset("CuentaContable")
                                    Else
                                      MsgBox "Incentivos no tienen Cuenta Contable", vbCritical, "Zeus Contabilidad"
                                      Exit Sub
                                    End If
                                     If ValidarCuentas(CtaIncentivos(i)) = False Then Print #1, CtaIncentivos(i); ExisteCodigo = False
                                     MontoIncentivos(i) = Me.AdoIcentivos.Recordset("Valor")
                                     i = i + 1
                                     Me.AdoIcentivos.Recordset.MoveNext
                                    Loop
                                End If
                                
                                
                                '//////////////////////////////CARGO LOS MONTOS DE LA NOMINA /////////////////////////////////
                                MontoSueldos = Me.AdoProcesos.Recordset("SalarioBasico") + Me.AdoProcesos.Recordset("Destajo") + Me.AdoProcesos.Recordset("Comisiones")
                                MontoProvAguinaldo = Me.AdoProcesos.Recordset("Mes13")
                                MontoProvAguinaldo = Me.AdoProcesos.Recordset("Vacaciones")
                                MontoOtrosIngresos = Me.AdoProcesos.Recordset("OtrosIngresos")
                                MontoHorasExtra = Me.AdoProcesos.Recordset("HorasExtras")
                                MontoINSSPatronal = Me.AdoProcesos.Recordset("INSSPatronal")
                                MontoINATEC = Me.AdoProcesos.Recordset("INATEC")
                                 '//////////////////////////DEDUCCIONES /////////////////////////////////////////
                                MontoAguinaldoPagar = Me.AdoProcesos.Recordset("Mes13")
                                MontoVacacionesPagar = Me.AdoProcesos.Recordset("Vacaciones")
                                MontoINSSPagar = Me.AdoProcesos.Recordset("MontoINSS")
                                MontoINATECPagar = Me.AdoProcesos.Recordset("INATEC")
                                MontoIRPagar = Me.AdoProcesos.Recordset("MontoIR")
                                MontoPrestamoPagar = Me.AdoProcesos.Recordset("Prestamo")
                                MontoNominaPagar = Me.AdoProcesos.Recordset("Salario") - Me.AdoProcesos.Recordset("DeduccionesTotal")
                                
                                
                                '///////////////////////////////////////////////////////////////////////////////////////////////////////
                                '///////////////////////BUSCO LA NOMINA DE SUBSIDIO ////////////////////////////////////////////////////
                                '//////////////////////////////////////////////////////////////////////////////////////////////////////
                                
                                
                                
                                If Reg = 1 Then
                                   '////////////////////////////////////////////////////////////////
                                   '////////AGREGO LOS INDICES DE TRANSACCIONES//////
                                   '///////////////////////////////////////////////////////////////
'                                    Resultado = GrabaEncabezado(NumeroPeriodo, NumeroTransaccion, Format(Me.DTPicker5.Value, "yyyy-mm-dd"), "Movimiento de Nominas", "Nominas", MonedaNomina)
                                    Reg = 2
                                End If
                                
                                
                                   '////////////////////////////////////////////////////////////////
                                   '////////AGREGO EL DETALLE DE LA TRANSACCION//////
                                   '///////////////////////////////////////////////////////////////
'                                    Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, Debito, Credito, "VTAS", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                              
                               Me.osProgress1.Value = Me.osProgress1.Value + 1
                               Me.AdoProcesos.Recordset.MoveNext
                             Loop
                         
                         
                         
                   Me.AdoNominas.Recordset.MoveNext
                Loop
     Close #1
     
     
     '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
     '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
     '//////////////////////////////////////////CREO LOS MOVIMIENTOS CONTABLES /////////////////////////////////////////////////////////
     '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
     '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
     
     
      If ExisteCodigo = False Then
           MsgBox "No existen Cuentas", vbCritical, "Sistema Contable"
        
           Abrir = "notepad.exe " & Directorio
           Shell Abrir
           Exit Sub
      Else
      
                Select Case TipoFactura
                   Case "Nominas"
'                     SqlString = "SELECT Nomina.NumNomina, TipoNomina.Nomina, Nomina.FechaNomina, Nomina.Marca, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.OtrosIngresos, DetalleNomina.Incentivos, DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones AS Salario, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.INATEC, DetalleNomina.Mes13, DetalleNomina.Vacaciones, DetalleNomina.Deducciones + DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR AS DeduccionesTotal FROM Nomina INNER JOIN TipoNomina ON Nomina.CodTipoNomina = TipoNomina.CodTipoNomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina  " & _
'                                                  "WHERE (Nomina.CodTipoNomina = '" & CodTipoNomina & "') AND (Nomina.Contabilizado = 0)"
                     SqlString = "SELECT Nomina.NumNomina, TipoNomina.Nomina, Nomina.FechaNomina, Nomina.Marca, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.OtrosIngresos, DetalleNomina.Incentivos, DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones AS Salario, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.INATEC, DetalleNomina.Mes13, DetalleNomina.Vacaciones, DetalleNomina.Deducciones + DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR AS DeduccionesTotal FROM Nomina INNER JOIN TipoNomina ON Nomina.CodTipoNomina = TipoNomina.CodTipoNomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina  " & _
                                                  "WHERE (Nomina.CodTipoNomina = '" & CodTipoNomina & "') AND (Nomina.Contabilizado = 0) AND (Nomina.Marca = 1)"
                     Me.AdoProcesos.RecordSource = SqlString
                
                    Me.AdoNominas.RecordSource = "SELECT  Nomina.NumNomina, TipoNomina.Nomina, Nomina.FechaNominaINI, Nomina.FechaNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.Marca FROM  Nomina INNER JOIN TipoNomina ON Nomina.CodTipoNomina = TipoNomina.CodTipoNomina " & _
                              "WHERE (Nomina.CodTipoNomina = '" & CodTipoNomina & "') AND (Nomina.Contabilizado = 0) AND (Nomina.Marca = 1)"
                    Me.AdoNominas.Refresh
                
                End Select
      
           '/////////////////////////////////////////////////////////////////////////////////////////////////////////////
           '///////////////////////////////////////BUSCO EL PERIODO DEL MOVIMIENTO ///////////////////////////////////////
           '//////////////////////////////////////////////////////////////////////////////////////////////////////////////
                         
            Reg = 1
            mes = Month(Me.DTPicker5.Value)
            Año = Year(Me.DTPicker5.Value)
            FechaIni = CDate("1/" & Month(Me.DTPicker5.Value) & "/" & Year(Me.DTPicker5.Value))
            FechaFin = DateSerial(Año, mes + 1, 1 - 1)
            TasaCambio = 1
                 
            Me.AdoConsultas.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "'))"
            Me.AdoConsultas.Refresh
            If Not Me.AdoConsultas.Recordset.EOF Then
               Periodo = Me.AdoConsultas.Recordset("Periodo")
               NumeroPeriodo = Me.AdoConsultas.Recordset("NPeriodo")
               EstadoPeriodo = Me.AdoConsultas.Recordset("EstadoPeriodo")
                      
               If EstadoPeriodo <> "A" Then
                  MsgBox "Periodo esta Bloqueado o Cerrado", vbCritical, "Zeus Contable"
                  Exit Sub
               End If

               Me.AdoConsultas.Recordset("NTransacciones") = Me.AdoConsultas.Recordset("NTransacciones") + 1
               Me.AdoConsultas.Recordset.Update
               NumeroTransaccion = Me.AdoConsultas.Recordset("NTransacciones")
               If Reg = 1 Then
                  '////////////////////////////////////////////////////////////////
                  '////////AGREGO LOS INDICES DE TRANSACCIONES//////
                  '///////////////////////////////////////////////////////////////
                  MonedaNomina = "Córdobas"
                   Resultado = GrabaEncabezado(NumeroPeriodo, NumeroTransaccion, Format(Me.DTPicker5.Value, "yyyy-mm-dd"), "Movimiento de Nominas", "Nominas", MonedaNomina)
                   Reg = 2
               End If
               
               
              Me.AdoNominas.Refresh
              Do While Not Me.AdoNominas.Recordset.EOF

                         NumNomina = Me.AdoNominas.Recordset("NumNomina")
                             
                         
                             Me.AdoProcesos.RecordSource = SqlString
                             Me.AdoProcesos.Refresh
                             Me.osProgress1.Visible = True
                             Me.osProgress1.Min = 0
                             Me.osProgress1.Value = 0
                             If Not Me.AdoProcesos.Recordset.EOF Then
                             Me.AdoProcesos.Recordset.MoveFirst
                             Me.osProgress1.Max = Me.AdoProcesos.Recordset.RecordCount
                             End If
                             Me.AdoProcesos.Refresh
                             
                             
                             Do While Not Me.AdoProcesos.Recordset.EOF
                             
                                CodEmpleado = Me.AdoProcesos.Recordset("Codempleado")
                                
                                '//////////////////////////////////////////////////////////////////////////////////////////
                                '///////////////////////////////////BUSCO LAS CUENTAS CONTABLES ///////////////////////////
                                '/////////////////////////////////////////////////////////////////////////////////////////
                                SqlString = "SELECT  * From Historico Where (CodEmpleado = " & CodEmpleado & ")"
                                Me.AdoCuentas.RecordSource = SqlString
                                Me.AdoCuentas.Refresh
                                If Not Me.AdoCuentas.Recordset.EOF Then
                                    If Not IsNull(Me.AdoCuentas.Recordset("CuentaSueldos")) Then CtaSueldos = Me.AdoCuentas.Recordset("CuentaSueldos")
                                    If Not IsNull(Me.AdoCuentas.Recordset("ProvAguinaldo")) Then CtaProvAguinaldo = Me.AdoCuentas.Recordset("ProvAguinaldo")
                                    If Not IsNull(Me.AdoCuentas.Recordset("ProvVacaciones")) Then CtaProvVacaciones = Me.AdoCuentas.Recordset("ProvVacaciones")
                                    If Not IsNull(Me.AdoCuentas.Recordset("CuentaOtrosIngresos")) Then CtaOtrosIngresos = Me.AdoCuentas.Recordset("CuentaOtrosIngresos")
                                    If Not IsNull(Me.AdoCuentas.Recordset("CuentaHorasExtra")) Then CtaHorasExtra = Me.AdoCuentas.Recordset("CuentaHorasExtra")
                                    If Not IsNull(Me.AdoCuentas.Recordset("INSSPatronal")) Then CtaINSSPatronal = Me.AdoCuentas.Recordset("INSSPatronal")
                                    If Not IsNull(Me.AdoCuentas.Recordset("INATEC")) Then CtaINATEC = Me.AdoCuentas.Recordset("INATEC")
                                    If Not IsNull(Me.AdoCuentas.Recordset("AguinaldoxPagar")) Then CtaAguinaldoPagar = Me.AdoCuentas.Recordset("AguinaldoxPagar")
                                    If Not IsNull(Me.AdoCuentas.Recordset("VacacionesxPagar")) Then CtaVacacionesPagar = Me.AdoCuentas.Recordset("VacacionesxPagar")
                                    If Not IsNull(Me.AdoCuentas.Recordset("INSSxPagar")) Then CtaINSSPagar = Me.AdoCuentas.Recordset("INSSxPagar")
                                    If Not IsNull(Me.AdoCuentas.Recordset("INATECxPagar")) Then CtaINATECPagar = Me.AdoCuentas.Recordset("INATECxPagar")
                                    If Not IsNull(Me.AdoCuentas.Recordset("IRxPagar")) Then CtaIRPagar = Me.AdoCuentas.Recordset("IRxPagar")
                                    If Not IsNull(Me.AdoCuentas.Recordset("PrestamoxPagar")) Then CtaPrestamoPagar = Me.AdoCuentas.Recordset("PrestamoxPagar")
                                    If Not IsNull(Me.AdoCuentas.Recordset("NominaxPagar")) Then CtaNominaPagar = Me.AdoCuentas.Recordset("NominaxPagar")
                                    If Not IsNull(Me.AdoCuentas.Recordset("INSSPatronalPagar")) Then CtaInssPatronalPagar = Me.AdoCuentas.Recordset("INSSPatronalPagar")
                                    If Not IsNull(Me.AdoCuentas.Recordset("CuentaSubsidio")) Then CtaSubsidio = Me.AdoCuentas.Recordset("CuentaSubsidio")
                                End If
                                
                                
                                '//////////////////////////////CARGO LOS MONTOS DE LA NOMINA /////////////////////////////////
                                MontoSueldos = CDbl(Format(Me.AdoProcesos.Recordset("SalarioBasico"), "##,##0.00")) + CDbl(Format(Me.AdoProcesos.Recordset("Destajo"), "##,##0.00")) + CDbl(Format(Me.AdoProcesos.Recordset("Comisiones"), "##,##0.00"))
                                MontoProvAguinaldo = Format(Me.AdoProcesos.Recordset("Mes13"), "##,##0.00")
                                MontoProvVacaciones = Format(Me.AdoProcesos.Recordset("Vacaciones"), "##,##0.00")
                                MontoOtrosIngresos = Format(Me.AdoProcesos.Recordset("OtrosIngresos"), "##,##0.00")
                                MontoHorasExtra = Format(Me.AdoProcesos.Recordset("HorasExtras"), "##,##0.00")
                                MontoINSSPatronal = Format(Me.AdoProcesos.Recordset("INSSPatronal"), "##,##0.00")
                                MontoINATEC = Format(Me.AdoProcesos.Recordset("INATEC"), "##,##0.00")
                                 '//////////////////////////DEDUCCIONES /////////////////////////////////////////
                                MontoAguinaldoPagar = Format(Me.AdoProcesos.Recordset("Mes13"), "##,##0.00")
                                MontoVacacionesPagar = Format(Me.AdoProcesos.Recordset("Vacaciones"), "##,##0.00")
                                MontoINSSPagar = Format(Me.AdoProcesos.Recordset("MontoINSS"), "##,##0.00")
                                MontoINATECPagar = Format(Me.AdoProcesos.Recordset("INATEC"), "##,##0.00")
                                MontoIRPagar = Format(Me.AdoProcesos.Recordset("MontoIR"), "##,##0.00")
                                MontoPrestamoPagar = Format(Me.AdoProcesos.Recordset("Prestamo"), "##,##0.00")
                                MontoNominaPagar = MontoSueldos + MontoOtrosIngresos + MontoHorasExtra + CDbl(Format(Me.AdoProcesos.Recordset("Incentivos"), "##,##0.00")) - MontoINSSPagar - MontoIRPagar - MontoPrestamoPagar - CDbl(Format(Me.AdoProcesos.Recordset("Deducciones"), "##,##0.00"))
                                MontoInssPatronalPagar = Me.AdoProcesos.Recordset("INSSPatronal")
                                
                                Me.AdoBuscaNomina.RecordSource = "SELECT  * FROM  Nomina INNER JOIN  TipoNomina ON Nomina.CodTipoNomina = TipoNomina.CodTipoNomina Where (Nomina.NumNomina = " & NumNomina & ")"
                                Me.AdoBuscaNomina.Refresh
                                If Not Me.AdoBuscaNomina.Recordset.EOF Then
                                  DescripcionMovimiento = "Registrando Nomina " & Me.AdoBuscaNomina.Recordset("Nomina") & " No " & NumNomina & "   Desde " & Me.AdoBuscaNomina.Recordset("FechaNominaINI") & " Hasta " & Me.AdoBuscaNomina.Recordset("FechaNomina")
                                Else
                                  DescripcionMovimiento = "Registrando Nominas No " & NumNomina
                                End If
                                
                                
                                '///////////////////////////CREO LOS REGISTROS CONTABLES/////////////////////////////////
                                Credito = 0
                                Resultado = GrabaDetalleNomina(CtaSueldos, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, MontoSueldos, Credito, "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                Resultado = GrabaDetalleNomina(CtaProvAguinaldo, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, MontoProvAguinaldo, Credito, "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                Resultado = GrabaDetalleNomina(CtaProvVacaciones, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, MontoProvVacaciones, Credito, "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                Resultado = GrabaDetalleNomina(CtaOtrosIngresos, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, MontoOtrosIngresos, Credito, "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                Resultado = GrabaDetalleNomina(CtaHorasExtra, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, MontoHorasExtra, Credito, "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                Resultado = GrabaDetalleNomina(CtaINSSPatronal, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, MontoINSSPatronal, Credito, "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                Resultado = GrabaDetalleNomina(CtaINATEC, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, MontoINATEC, Credito, "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                '////////////////////////////BUSCO LAS CUENTAS DE INCENTIVOS ///////////////////////////////////////
                                
                                If Me.AdoProcesos.Recordset("Incentivos") <> 0 Then
                                    Me.AdoIcentivos.RecordSource = "SELECT * FROM DetalleIncentivo INNER JOIN Incentivo ON DetalleIncentivo.NumIncentivo = Incentivo.NumIncentivo INNER JOIN TipoIncentivo ON Incentivo.CodTipoIncentivo = TipoIncentivo.CodTipoIncentivo  " & _
                                                                   "Where (DetalleIncentivo.NumNomina = " & NumNomina & ") And (Incentivo.CodEmpleado = " & CodEmpleado & ") And (DetalleIncentivo.Valor <> 0)"
                                    Me.AdoIcentivos.Refresh
                                    i = 0
                                    Registros = Me.AdoIcentivos.Recordset.RecordCount
                                    ReDim CtaIncentivos(Registros)
                                    ReDim MontoIncentivos(Registros)
                                    
                                    If Not Me.AdoIcentivos.Recordset.EOF Then
                                     Me.AdoIcentivos.Recordset.MoveFirst
                                    End If
                                    Do While Not Me.AdoIcentivos.Recordset.EOF
                                     CtaIncentivos(i) = Me.AdoIcentivos.Recordset("CuentaContable")
                                     MontoIncentivos(i) = Format(Me.AdoIcentivos.Recordset("Valor"), "##,##0.00")
                                     Credito = 0
                                     Debito = MontoIncentivos(i)
                                     Resultado = GrabaDetalleNomina(CtaIncentivos(i), Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, MontoIncentivos(i), Credito, "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                     i = i + 1
                                     Me.AdoIcentivos.Recordset.MoveNext
                                    Loop
                                End If
                                
                                
                                
                                
                                Debito = 0
                                Resultado = GrabaDetalleNomina(CtaAguinaldoPagar, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, MontoAguinaldoPagar, "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                Resultado = GrabaDetalleNomina(CtaVacacionesPagar, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, MontoVacacionesPagar, "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                Resultado = GrabaDetalleNomina(CtaINATECPagar, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, MontoINATECPagar, "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                Resultado = GrabaDetalleNomina(CtaInssPatronalPagar, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, MontoInssPatronalPagar, "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                Resultado = GrabaDetalleNomina(CtaINSSPagar, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, MontoINSSPagar, "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                
                                Resultado = GrabaDetalleNomina(CtaIRPagar, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, MontoIRPagar, "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                Resultado = GrabaDetalleNomina(CtaPrestamoPagar, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, MontoPrestamoPagar, "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                
                                
                                '////////////////////////////BUSCO LAS CUENTAS DE DEDUCCIONES ///////////////////////////////////////
                                If Me.AdoProcesos.Recordset("Deducciones") <> 0 Then
                                    Me.AdoDeducciones.RecordSource = "SELECT  DetalleDeduccion.*, TipoDeduccion.CuentaContable, TipoDeduccion.Deduccion, Deduccion.CodEmpleado FROM  DetalleDeduccion INNER JOIN Deduccion ON DetalleDeduccion.NumDeduccion = Deduccion.NumDeduccion INNER JOIN TipoDeduccion ON Deduccion.CodTipoDeduccion = TipoDeduccion.CodTipoDeduccion  " & _
                                                      "WHERE (DetalleDeduccion.NumNomina = " & NumNomina & ") AND (Deduccion.CodEmpleado = " & CodEmpleado & ") AND (DetalleDeduccion.Valor <> 0)"
                                    Me.AdoDeducciones.Refresh
                                    i = 0
                                    Registros = Me.AdoDeducciones.Recordset.RecordCount
                                    ReDim CtaDeducciones(Registros)
                                    ReDim MontoDeducciones(Registros)
                                    If Not Me.AdoDeducciones.Recordset.EOF Then
                                     Me.AdoDeducciones.Recordset.MoveFirst
                                    End If
                                    Do While Not Me.AdoDeducciones.Recordset.EOF
                                     CtaDeducciones(i) = Me.AdoDeducciones.Recordset("CuentaContable")
                                     MontoDeducciones(i) = Format(Me.AdoDeducciones.Recordset("Valor"), "##,##0.00")
                                     Debito = 0
                                     Resultado = GrabaDetalleNomina(CtaDeducciones(i), Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, MontoDeducciones(i), "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                     i = i + 1
                                     Me.AdoDeducciones.Recordset.MoveNext
                                    Loop
                                End If
                                
                               
                                Resultado = GrabaDetalleNomina(CtaNominaPagar, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, MontoNominaPagar, "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                
                                
                                 '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                '////////////////////////////////////////AHORA CONTABILIZO LA NOMINA DE SUBSIDIO ///////////////////////////////////////
                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                             
'                                Me.AdoNominaSubsidio.RecordSource = "SELECT  NumNominaSubsidio, CodEmpleado, SUM(Subsidio) AS Monto From DetalleNomSubsidio GROUP BY NumNominaSubsidio, CodEmpleado Having (NumNominaSubsidio = " & NumNomina & ") AND (CodEmpleado = " & CodEmpleado & ")"
'                                Me.AdoNominaSubsidio.RecordSource = "SELECT DetalleNomSubsidio.NumNominaSubsidio, DetalleNomSubsidio.CodEmpleado, DetalleNomSubsidio.Subsidio As Monto, Subsidio.NumSubsidio, Subsidio.CodEmpleado AS Expr1, Subsidio.CodTipoSubsidio FROM DetalleNomSubsidio INNER JOIN Subsidio ON DetalleNomSubsidio.NumNominaSubsidio = Subsidio.NumSubsidio AND DetalleNomSubsidio.CodEmpleado = Subsidio.CodEmpleado Where (DetalleNomSubsidio.CodEmpleado = " & CodEmpleado & ") And (DetalleNomSubsidio.NumNominaSubsidio = " & NumNomina & ")"
                                Me.AdoNominaSubsidio.RecordSource = "SELECT  Empleado.CodEmpleado1, Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Subsidio.NumSubsidio, TipoSubsidio.Subsidio, DetalleSubsidio.Valor AS Monto, DetalleSubsidio.NumVez, DetalleSubsidio.Descripcion, DetalleSubsidio.NumNominaSubsidio, TipoSubsidio.CodTipoSubsidio FROM TipoSubsidio INNER JOIN Empleado INNER JOIN Subsidio ON Empleado.CodEmpleado = Subsidio.CodEmpleado INNER JOIN DetalleSubsidio ON Subsidio.NumSubsidio = DetalleSubsidio.NumSubsidio ON TipoSubsidio.CodTipoSubsidio = Subsidio.CodTipoSubsidio Where (DetalleSubsidio.NumNominaSubsidio = " & NumNomina & ") And (Empleado.CodEmpleado = " & CodEmpleado & ") ORDER BY Empleado.CodEmpleado"
                                Me.AdoNominaSubsidio.Refresh
                                Do While Not Me.AdoNominaSubsidio.Recordset.EOF
                                  '------------------BUSCO LA CUENTA DESUBSIDIO ----------------------------------
                                  CodSubsidio = Me.AdoNominaSubsidio.Recordset("CodTipoSubsidio")
                                  Me.AdoConsultaNomina.RecordSource = "SELECT  Empleado.CodEmpleado, CuentasIncentivos.CodIncentivo, CuentasIncentivos.CodCuentas, Empleado.CodEmpleado1 FROM CuentasIncentivos INNER JOIN Empleado ON CuentasIncentivos.CodEmpleado = Empleado.CodEmpleado1 WHERE (Empleado.CodEmpleado = " & CodEmpleado & ") AND (CuentasIncentivos.CodIncentivo = '" & CodSubsidio & "')"
                                  Me.AdoConsultaNomina.Refresh
                                  If Not Me.AdoConsultaNomina.Recordset.EOF Then
                                    CtaSubsidio = Me.AdoConsultaNomina.Recordset("CodCuentas")
                                  End If
                                  
                                  
                                  Debito = Me.AdoNominaSubsidio.Recordset("Monto")
                                  Credito = 0
                                  Resultado = GrabaDetalleNomina(CtaSubsidio, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento & "Empleado: " & CodEmpleado, "Debito", TasaCambio, Debito, Credito, "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                  
                                  Debito = 0
                                  Credito = Me.AdoNominaSubsidio.Recordset("Monto")
                                  Resultado = GrabaDetalleNomina(CtaNominaPagar, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, Credito, "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                  Me.AdoNominaSubsidio.Recordset.MoveNext
                                Loop
                                
                                   '////////////////////////////////////////////////////////////////
                                   '////////AGREGO EL DETALLE DE LA TRANSACCION//////
                                   '///////////////////////////////////////////////////////////////
'                                    Resultado = GrabaDetalleNomina(CodigoCuentaInventario, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, Debito, Credito, "VTAS", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                              
                               Me.osProgress1.Value = Me.osProgress1.Value + 1
                               Me.AdoProcesos.Recordset.MoveNext
                             Loop
                             
                             
                                 '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                '////////////////////////////////////////AHORA CONTABILIZO LA NOMINA DE SUBSIDIO ///////////////////////////////////////
                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                             
'                                Me.AdoNominaSubsidio.RecordSource = "SELECT  NumNominaSubsidio, CodEmpleado, SUM(Subsidio) AS Monto From DetalleNomSubsidio GROUP BY NumNominaSubsidio, CodEmpleado Having (NumNominaSubsidio = " & NumNomina & ")"
'                                Me.AdoNominaSubsidio.Refresh
'                                Do While Not Me.AdoNominaSubsidio.Recordset.EOF
'                                  Debito = Me.AdoNominaSubsidio.Recordset("Monto")
'                                  Credito = 0
'                                  Resultado = GrabaDetalleNomina(CtaSubsidio, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento & "Empleado: " & CodEmpleado, "Debito", TasaCambio, Debito, Credito, "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
'                                  Me.AdoNominaSubsidio.Recordset.MoveNext
'                                Loop
                    
'                                Me.AdoNominaSubsidio.RecordSource = "SELECT  NumNominaSubsidio, MAX(CodEmpleado) AS Expr1, SUM(Subsidio) AS Monto From DetalleNomSubsidio GROUP BY NumNominaSubsidio Having (NumNominaSubsidio = " & NumNomina & ") ORDER BY MAX(CodEmpleado)"
'                                Me.AdoNominaSubsidio.Refresh
'                                If Not Me.AdoNominaSubsidio.Recordset.EOF Then
'                                   Debito = 0
'                                   Credito = Me.AdoNominaSubsidio.Recordset("Monto")
'                                   Resultado = GrabaDetalleNomina(CtaNominaPagar, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, Credito, "Nominas", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
'
'                                End If
                             
                             
                             
                    Me.AdoNominas.Recordset.MoveNext
                Loop
               
               

            
               
               
               
            End If
      
      End If
      
       
      
      
      
      '///////////////////////////////////////////ACTUALIZO LA NOMINA ////////////////////////////////////////////////////////////
       Me.AdoConsultaNomina.RecordSource = "SELECT  * From Nomina WHERE (NumNomina = " & NumNomina & ")"
       Me.AdoConsultaNomina.Refresh
       If Not Me.AdoConsultaNomina.Recordset.EOF Then
          Me.AdoConsultaNomina.Recordset("Contabilizado") = True
          Me.AdoConsultaNomina.Recordset.Update
          CmdConsultar_Click
       End If


   '----------------------------------------------------------------------------------------------------------------------
   '---------------------------------SI MARCAN EL CHECK GENERO UN CHEQUE PARA CADA EMPLEADO-------------------------------
   '-----------------------------------------------------------------------------------------------------------------------
   If Me.ChkCheques.Value = xtpChecked Then
   
      Select Case TipoFactura
           Case "Nominas"
                     SqlString = "SELECT Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.*, Empleado.* FROM DetalleNomina INNER JOIN Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado  " & _
                                  "Where (DetalleNomina.NumNomina = " & NumNomina & ")"
                  
      End Select
   
      Me.AdoEmpleados.RecordSource = SqlString
      Me.AdoEmpleados.Refresh
      
      
      Me.osProgress1.Visible = True
      Me.osProgress1.Min = 0
      Me.osProgress1.Value = 0
      If Not Me.AdoEmpleados.Recordset.EOF Then
         Me.AdoEmpleados.Recordset.MoveFirst
         Me.osProgress1.Max = Me.AdoEmpleados.Recordset.RecordCount
      End If
      
      Me.AdoEmpleados.Recordset.MoveFirst
      Do While Not Me.AdoEmpleados.Recordset.EOF
      
           CodEmpleado = Me.AdoEmpleados.Recordset("Codempleado")
                                
                                '//////////////////////////////////////////////////////////////////////////////////////////
                                '///////////////////////////////////BUSCO LAS CUENTAS CONTABLES ///////////////////////////
                                '/////////////////////////////////////////////////////////////////////////////////////////
                                SqlString = "SELECT  * From Historico Where (CodEmpleado = " & CodEmpleado & ")"
                                Me.AdoCuentas.RecordSource = SqlString
                                Me.AdoCuentas.Refresh
                                If Not Me.AdoCuentas.Recordset.EOF Then
                                If Not IsNull(Me.AdoCuentas.Recordset("CuentaSueldos")) Then
                                    CtaSueldos = Me.AdoCuentas.Recordset("CuentaSueldos")
                                End If
                                    CtaProvAguinaldo = Me.AdoCuentas.Recordset("ProvAguinaldo")
                                    CtaProvVacaciones = Me.AdoCuentas.Recordset("ProvVacaciones")
                                    CtaOtrosIngresos = Me.AdoCuentas.Recordset("CuentaOtrosIngresos")
                                    CtaHorasExtra = Me.AdoCuentas.Recordset("CuentaHorasExtra")
                                    CtaINSSPatronal = Me.AdoCuentas.Recordset("INSSPatronal")
                                    CtaINATEC = Me.AdoCuentas.Recordset("INATEC")
                                    CtaAguinaldoPagar = Me.AdoCuentas.Recordset("AguinaldoxPagar")
                                    CtaVacacionesPagar = Me.AdoCuentas.Recordset("VacacionesxPagar")
                                    CtaINSSPagar = Me.AdoCuentas.Recordset("INSSxPagar")
                                    CtaINATECPagar = Me.AdoCuentas.Recordset("INATECxPagar")
                                    CtaIRPagar = Me.AdoCuentas.Recordset("IRxPagar")
                                    CtaPrestamoPagar = Me.AdoCuentas.Recordset("PrestamoxPagar")
                                    CtaNominaPagar = Me.AdoCuentas.Recordset("NominaxPagar")
                                    CtaInssPatronalPagar = Me.AdoCuentas.Recordset("INSSPatronalPagar")
                                    CtaSubsidio = Me.AdoCuentas.Recordset("CuentaSubsidio")
                                    CtaBancos = Me.AdoCuentas.Recordset("CuentaBanco")
                                End If

        
            mes = Month(Me.DTPicker5.Value)
            Año = Year(Me.DTPicker5.Value)
            FechaIni = CDate("1/" & Month(Me.DTPicker5.Value) & "/" & Year(Me.DTPicker5.Value))
            FechaFin = DateSerial(Año, mes + 1, 1 - 1)
            TasaCambio = 1
                 
            Me.AdoConsultas.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "'))"
            Me.AdoConsultas.Refresh
            If Not Me.AdoConsultas.Recordset.EOF Then
               Periodo = Me.AdoConsultas.Recordset("Periodo")
               NumeroPeriodo = Me.AdoConsultas.Recordset("NPeriodo")
               EstadoPeriodo = Me.AdoConsultas.Recordset("EstadoPeriodo")
                      
               If EstadoPeriodo <> "A" Then
                  MsgBox "Periodo esta Bloqueado o Cerrado", vbCritical, "Zeus Contable"
                  Exit Sub
               End If
               
               
               '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
               '////////////////////////////////////BUSCO EL MONTO A PAGAR ///////////////////////////////////////////////////
               '////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                MontoSueldos = CDbl(Format(Val(Me.AdoEmpleados.Recordset("SalarioBasico")), "##,##0.00")) + CDbl(Format(Val(Me.AdoEmpleados.Recordset("Destajo")), "##,##0.00")) + CDbl(Format(Val(Me.AdoEmpleados.Recordset("Comisiones")), "##,##0.00"))
                                MontoOtrosIngresos = Format(Val(Me.AdoEmpleados.Recordset("OtrosIngresos")), "##,##0.00")
                                MontoHorasExtra = Format(Val(Me.AdoEmpleados.Recordset("HorasExtras")), "##,##0.00")
                                 '//////////////////////////DEDUCCIONES /////////////////////////////////////////
                                MontoINSSPagar = Format(Val(Me.AdoEmpleados.Recordset("MontoINSS")), "##,##0.00")
                                MontoIRPagar = Format(Val(Me.AdoEmpleados.Recordset("MontoIR")), "##,##0.00")
                                MontoPrestamoPagar = Format(Val(Me.AdoEmpleados.Recordset("Prestamo")), "##,##0.00")
                               
             
                                '-------SUBSIDIOS --------------------------
                                Me.AdoNominaSubsidio.RecordSource = "SELECT  NumNominaSubsidio, CodEmpleado, SUM(Subsidio) AS Monto From DetalleNomSubsidio GROUP BY NumNominaSubsidio, CodEmpleado Having (NumNominaSubsidio = " & NumNomina & ") AND (CodEmpleado = " & CodEmpleado & ")"
                                Me.AdoNominaSubsidio.Refresh
                                If Not Me.AdoNominaSubsidio.Recordset.EOF Then
                                   MontoSubsidio = Format(Val(Me.AdoNominaSubsidio.Recordset("Monto")), "##,##0.00")
                                Else
                                   MontoSubsidio = 0
                                End If
                                
                                 MontoNominaPagar = MontoSubsidio + MontoSueldos + MontoOtrosIngresos + MontoHorasExtra + CDbl(Format(Me.AdoEmpleados.Recordset("Incentivos"), "##,##0.00")) - MontoINSSPagar - MontoIRPagar - MontoPrestamoPagar - CDbl(Format(Me.AdoEmpleados.Recordset("Deducciones")))
               

               Me.AdoConsultas.Recordset("NTransacciones") = Me.AdoConsultas.Recordset("NTransacciones") + 1
               Me.AdoConsultas.Recordset.Update
               NumeroTransaccion = Me.AdoConsultas.Recordset("NTransacciones")

                  '////////////////////////////////////////////////////////////////
                  '////////AGREGO LOS INDICES DE TRANSACCIONES//////
                  '///////////////////////////////////////////////////////////////
                   MonedaNomina = "Córdobas"
                   Resultado = GrabaEncabezado(NumeroPeriodo, NumeroTransaccion, Format(Me.DTPicker5.Value, "yyyy-mm-dd"), "Pago de Nominas", "CHEQUE", MonedaNomina)
                   Reg = 2

             End If
             
              '//////////////////////////////////////CUENTA X PAGAR /////////////////////////////////////////////////////
              Credito = 0
              NumeroFactura = "-"
              Resultado = GrabaDetalleNomina(CtaNominaPagar, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, MontoNominaPagar, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
             
              
              '////////////////////////////////////CUENTA DE BANCO //////////////////////////////////////////////////////////////////
              '#######
              Debito = 0
              NumeroFactura = "#######"
              TipoFactura = Me.AdoEmpleados.Recordset("Nombres")
              Resultado = GrabaDetalleNomina(CtaBancos, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, MontoNominaPagar, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, TipoFactura)
        
        
        
        
        
        
        
        Me.osProgress1.Value = Me.osProgress1.Value + 1
        Me.AdoEmpleados.Recordset.MoveNext
      Loop
   End If

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    With Me.AdoDatosEmpresa
       .ConnectionString = Conexion
       .RecordSource = "SELECT  * From DatosEmpresa"
       .Refresh
    End With

    If Not IsNull(Me.AdoDatosEmpresa.Recordset("ConexionNomina")) Then
      ConexionNominas = Me.AdoDatosEmpresa.Recordset("ConexionNomina")
    Else
      MsgBox "No Existe conexion con la Nominas", vbCritical, "Zeus Contabilidad"
      Unload Me
    End If
    
   With Me.AdoNominaSubsidio
       .ConnectionString = ConexionNominas
    End With
    
   With Me.AdoBuscaNomina
       .ConnectionString = ConexionNominas
    End With
    
   With Me.AdoCuentas
       .ConnectionString = ConexionNominas
    End With
    
    With Me.AdoIcentivos
       .ConnectionString = ConexionNominas
    End With
    
    With Me.AdoDeducciones
       .ConnectionString = ConexionNominas
    End With
    
    With Me.AdoSubsidios
       .ConnectionString = ConexionNominas
    End With
    
    With Me.AdoDetalleNomina
       .ConnectionString = ConexionNominas
    End With
    
    
    With Me.AdoNominas
       .ConnectionString = ConexionNominas
    End With
    
    With Me.AdoTipoNominas
       .ConnectionString = ConexionNominas
       .RecordSource = "SELECT CodTipoNomina, Nomina, Periodo From TipoNomina"
       .Refresh
    End With
    
    With Me.AdoProcesos
       .ConnectionString = ConexionNominas
    End With

    With Me.AdoConsultas
       .ConnectionString = Conexion
    End With
    
    With Me.AdoConsultaNomina
       .ConnectionString = ConexionNominas
    End With

    With Me.AdoEmpleados
       .ConnectionString = ConexionNominas
    End With
End Sub

