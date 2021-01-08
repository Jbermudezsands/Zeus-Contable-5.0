VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmTransacciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Transacciones"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13635
   Icon            =   "FrmTransacciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   13635
   Begin VB.CheckBox ChkVentana 
      Caption         =   "Mostrar Vtana Factura"
      Height          =   255
      Left            =   360
      TabIndex        =   52
      Top             =   4320
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc AdoFuente 
      Height          =   375
      Left            =   9720
      Top             =   6960
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "AdoFuente"
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
   Begin VB.CommandButton CmdMemorizar 
      Caption         =   "Memorizar"
      Height          =   375
      Left            =   5040
      TabIndex        =   49
      Top             =   5040
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc AdoProveedores 
      Height          =   375
      Left            =   10200
      Top             =   7440
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "AdoProveedores"
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
   Begin MSAdodcLib.Adodc AdoBuscar 
      Height          =   330
      Left            =   7320
      Top             =   6360
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "AdoBuscar"
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
   Begin MSAdodcLib.Adodc AdoBuscaCuenta 
      Height          =   375
      Left            =   3840
      Top             =   6360
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
      Caption         =   "AdoBuscaCuenta"
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
   Begin TrueOleDBGrid80.TDBGrid TDBGridFechas1 
      Height          =   855
      Left            =   360
      TabIndex        =   26
      Top             =   3000
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1508
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
      Splits(0).DividerColor=   14215660
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
      AllowDelete     =   -1  'True
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
   Begin VB.PictureBox TDBGridFechas 
      Height          =   2055
      Left            =   4800
      ScaleHeight     =   1995
      ScaleWidth      =   6315
      TabIndex        =   27
      Top             =   1680
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CommandButton Command1 
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
         Left            =   3360
         Picture         =   "FrmTransacciones.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1320
         Width           =   375
      End
      Begin XtremeSuiteControls.CheckBox ChkFactura 
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   2160
         Visible         =   0   'False
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fac Principal"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   420
         Left            =   120
         TabIndex        =   45
         Top             =   760
         Width           =   3135
         _Version        =   786432
         _ExtentX        =   5530
         _ExtentY        =   741
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton OptFacturaCompra 
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   120
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Factura Compra"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton OptFacturaVenta 
            Height          =   255
            Left            =   1680
            TabIndex        =   47
            Top             =   120
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Factura Venta"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin VB.TextBox TxtProveedor 
         Height          =   375
         Left            =   1800
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   2280
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   4920
         TabIndex        =   43
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4920
         TabIndex        =   42
         Top             =   240
         Width           =   1095
      End
      Begin TrueOleDBList80.TDBCombo TDBProveedor 
         Bindings        =   "FrmTransacciones.frx":0458
         Height          =   315
         Left            =   1200
         TabIndex        =   40
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
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
         ListField       =   "CodCuentas"
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
         _PropDict       =   $"FrmTransacciones.frx":0475
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=12,.bold=0,.fontsize=825,.italic=0"
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
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   615
         Left            =   3360
         TabIndex        =   36
         Top             =   600
         Width           =   2535
         _Version        =   786432
         _ExtentX        =   4471
         _ExtentY        =   1085
         _StockProps     =   79
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton OptIva 
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   120
            Width           =   1095
            _Version        =   786432
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Causa IVA"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton OptRetencion 
            Height          =   375
            Left            =   1320
            TabIndex        =   38
            Top             =   120
            Width           =   1095
            _Version        =   786432
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Causa Retencion"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin MSComCtl2.DTPicker DTPFechaCredito 
         Height          =   300
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   80805889
         CurrentDate     =   38918
      End
      Begin VB.TextBox TxtMonto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPFechaVence 
         Height          =   300
         Left            =   3240
         TabIndex        =   28
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   80805889
         CurrentDate     =   38918
      End
      Begin VB.Label LblNombres 
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1560
         Width           =   4455
      End
      Begin VB.Label LblProveedor 
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Tipo de Factura"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   0
         TabIndex        =   35
         Top             =   550
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Fecha Vence"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   34
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Monto Desc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   33
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Fecha Descuento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   0
         Width           =   1695
      End
   End
   Begin MSAdodcLib.Adodc AdoFechasVence 
      Height          =   375
      Left            =   600
      Top             =   6480
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "AdoFechasVence"
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
      Left            =   8640
      TabIndex        =   25
      Top             =   5040
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   255
      Left            =   11520
      OleObjectBlob   =   "FrmTransacciones.frx":051F
      TabIndex        =   23
      Top             =   4560
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   10080
      OleObjectBlob   =   "FrmTransacciones.frx":058B
      TabIndex        =   22
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   12480
      TabIndex        =   16
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton SmartButton1 
      Caption         =   "Borrar Linea"
      Height          =   375
      Left            =   7440
      TabIndex        =   15
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   6240
      TabIndex        =   14
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton CmdSiguiente 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "Anterior"
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   5040
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   810
      ItemData        =   "FrmTransacciones.frx":05F5
      Left            =   960
      List            =   "FrmTransacciones.frx":05FF
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc DtaHistorial 
      Height          =   330
      Left            =   720
      Top             =   7920
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "DtaHistorial"
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
   Begin MSAdodcLib.Adodc DtaCuentas 
      Height          =   375
      Left            =   600
      Top             =   7440
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "DtaCuentas"
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
   Begin MSAdodcLib.Adodc DtaPeriodos 
      Height          =   375
      Left            =   600
      Top             =   6960
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "DtaPeriodos"
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
   Begin MSAdodcLib.Adodc DtaTransacciones 
      Height          =   375
      Left            =   6600
      Top             =   7560
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
      Caption         =   "DtaTransacciones"
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
   Begin MSAdodcLib.Adodc DtaTransaccionesNuevas 
      Height          =   375
      Left            =   6600
      Top             =   7200
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
      Caption         =   "DtaTransaccionesNuevas"
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
   Begin MSAdodcLib.Adodc DtaIndice 
      Height          =   375
      Left            =   6600
      Top             =   6840
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
      Caption         =   "DtaIndice"
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
      Left            =   3600
      Top             =   7680
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSAdodcLib.Adodc DtaNacceso 
      Height          =   375
      Left            =   3600
      Top             =   7200
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "DtaNacceso"
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
   Begin MSAdodcLib.Adodc DtaTasas 
      Height          =   375
      Left            =   3600
      Top             =   6840
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "DtaTasas"
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
   Begin VB.CommandButton CmdBuscarEmpleado 
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
      Left            =   6840
      Picture         =   "FrmTransacciones.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox TxtDiferencia 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox TxtDebito 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox TxtCredito 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13335
      Begin MSDataListLib.DataCombo TxtFuente 
         Bindings        =   "FrmTransacciones.frx":0762
         DataSource      =   "AdoFuente"
         Height          =   315
         Left            =   11640
         TabIndex        =   50
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Fuente"
         BoundColumn     =   "Fuente"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker TxtFecha 
         Height          =   285
         Left            =   720
         TabIndex        =   31
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Format          =   80805889
         CurrentDate     =   38918
      End
      Begin VB.ComboBox CmbMoneda 
         Height          =   315
         ItemData        =   "FrmTransacciones.frx":0789
         Left            =   8520
         List            =   "FrmTransacciones.frx":0793
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TxtNTransacciones 
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "0"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtFuentess 
         Height          =   285
         Left            =   11640
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtPeriodo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   7440
         OleObjectBlob   =   "FrmTransacciones.frx":07AA
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmTransacciones.frx":081E
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "FrmTransacciones.frx":0886
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   2760
         OleObjectBlob   =   "FrmTransacciones.frx":0902
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   11040
         OleObjectBlob   =   "FrmTransacciones.frx":096E
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   255
      Left            =   2880
      OleObjectBlob   =   "FrmTransacciones.frx":09D8
      TabIndex        =   24
      Top             =   4080
      Width           =   1335
   End
   Begin TrueOleDBGrid80.TDBGrid DBGTransacciones 
      Bindings        =   "FrmTransacciones.frx":0A4A
      Height          =   3015
      Left            =   120
      TabIndex        =   53
      Top             =   960
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   5318
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
Attribute VB_Name = "FrmTransacciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Dim RstFechas As ADODB.Recordset

Private Sub CmbTipoFactura_Change()


End Sub

Private Sub CmdAceptar_Click()

On Error GoTo TipoErrs:

  TDBGridFechas.Visible = False
If Me.TDBProveedor.Text = "" Then
  MsgBox "Necesita Seleccionar un Proveedor", vbCritical, "Sistema Contable"
  Exit Sub
End If

FechaFactura = Format(Me.DTPFechaCredito.Value, "dd/mm/yyyy")
FechaVence = Format(Me.DTPFechaVence.Value, "dd/mm/yyyy")
Monto = Val(Me.TxtMonto.Text)


Me.DBGTransacciones.Columns(17).Text = Format(FechaFactura, "dd/mm/yyyy")
Me.DBGTransacciones.Columns(18).Text = Monto
Me.DBGTransacciones.Columns(19).Text = Format(FechaVence, "dd/mm/yyyy")

Me.DBGTransacciones.Columns(20).Text = Me.TDBProveedor.Text
If Me.OptFacturaCompra.Value = True Then
  Me.DBGTransacciones.Columns(21).Text = "FacturaCompra"
Else
  Me.DBGTransacciones.Columns(21).Text = "FacturaVenta"
End If

Me.DBGTransacciones.Enabled = True
 
     FrmTransacciones.DBGTransacciones.Columns(0).Button = True
     FrmTransacciones.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
     FrmTransacciones.DBGTransacciones.Columns(6).Button = True
     FrmTransacciones.DBGTransacciones.Columns(6).Locked = True
     FrmTransacciones.DBGTransacciones.Columns(0).Width = 1500
     FrmTransacciones.DBGTransacciones.Columns(2).Width = 1000
     FrmTransacciones.DBGTransacciones.Columns(3).Caption = "Descripcion"
     FrmTransacciones.DBGTransacciones.Columns(4).Width = 1000
     FrmTransacciones.DBGTransacciones.Columns(4).Button = True
     FrmTransacciones.DBGTransacciones.Columns(5).Width = 1000
     FrmTransacciones.DBGTransacciones.Columns(6).Width = 800
     FrmTransacciones.DBGTransacciones.Columns(7).Width = 1200
     FrmTransacciones.DBGTransacciones.Columns(7).NumberFormat = "##,##0.000000"
     FrmTransacciones.DBGTransacciones.Columns(8).Width = 1200
     FrmTransacciones.DBGTransacciones.Columns(8).NumberFormat = "##,##0.00"
     FrmTransacciones.DBGTransacciones.Columns(9).Width = 1200
     FrmTransacciones.DBGTransacciones.Columns(9).NumberFormat = "##,##0.00"
     FrmTransacciones.DBGTransacciones.Columns(10).Visible = False
     FrmTransacciones.DBGTransacciones.Columns(11).Visible = False
     FrmTransacciones.DBGTransacciones.Columns(12).Visible = False
     FrmTransacciones.DBGTransacciones.Columns(13).Visible = False
     FrmTransacciones.DBGTransacciones.Columns(14).Visible = False
     FrmTransacciones.DBGTransacciones.Columns(15).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(16).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(17).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(18).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(19).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(20).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(21).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(22).Visible = False
 
      FrmTransacciones.DBGTransacciones.SetFocus
      FrmTransacciones.DBGTransacciones.PostMsg (5)
 
 Exit Sub
TipoErrs:
  MsgBox err.Description
End Sub

Private Sub CmdAnterior_Click()
On Error GoTo TipoErrs
 If Not Val(Me.TxtDiferencia.Text) = 0 Then
   MsgBox "El Movimiento esta Desbalanceado", vbCritical, "Sistema Contable"
   Exit Sub
  End If
 Primero = True
    
  Me.CmbMoneda.Enabled = False
  '//////Grabo las descripcion en los indices//////////////////////
 Me.DBGTransacciones.Enabled = True
 mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente, IndiceTransaccion.TipoMoneda From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between '" & Format(NumFecha1, "yyyymmdd") & "' And '" & Format(NumFecha2, "yyyymmdd") & "' ) AND ((IndiceTransaccion.NumeroMovimiento)= " & NumeroTransaccion & "))"
 Me.DtaConsulta.Refresh
 
 If Not DtaTransacciones.Recordset.EOF Then
 Me.DtaTransacciones.Recordset.MoveFirst
 End If
       
       If Not DtaConsulta.Recordset.EOF Then
         If Not IsNull(Me.CmbMoneda.Text = Me.DtaConsulta.Recordset("TipoMoneda")) Then
          'Me.CmbMoneda.Text = Me.DtaConsulta.Recordset("TipoMoneda")
         End If
        If Not Me.DBGTransacciones.Columns(3).Text = "" Then
 
          'Me.'DtaConsulta.Recordset.Edit
'          Me.DtaConsulta.Recordset("TipoMoneda") = Me.CmbMoneda.Text
'          Me.DtaConsulta.Recordset("DescripcionMovimiento") = Me.DBGTransacciones.Columns(3).Text
'          Me.DtaConsulta.Recordset.Update
        Else
          'Me.'DtaConsulta.Recordset.Edit
          'Me.DtaConsulta.Recordset("TipoMoneda") = Me.CmbMoneda.Text
'          Me.DtaConsulta.Recordset.Update
        End If
       End If
 TotalDiferencia = 0
 TotalCredito = 0
 TotalDebito = 0
 Debito = 0
 Credito = 0
 Diferencia = 0
 
If Me.TxtNTransacciones = 0 Then
 mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones WHERE (((Transacciones.NombreCuenta)<>'**********CANCELADO*************') AND ((Transacciones.FechaTransaccion) Between '" & Format(NumFecha1, "yyyymmdd") & "' And '" & Format(NumFecha2, "yyyymmdd") & "')) ORDER BY Transacciones.NumeroMovimiento"
 'Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & "))ORDER BY Transacciones.NumeroMovimiento"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
   NumeroTransaccion = DtaConsulta.Recordset("NumeroMovimiento")
   Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.NombreCuenta)<>'**********CANCELADO*************') AND ((Transacciones.FechaTransaccion) Between '" & Format(NumFecha1, "yyyymmdd") & "' And '" & Format(NumFecha2, "yyyymmdd") & "') AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion "
   'Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion"
   Me.DtaTransacciones.Refresh
   If Not DtaTransacciones.Recordset.EOF Then
     Me.TxtFecha.Value = Me.DtaTransacciones.Recordset("FechaTransaccion")
     Me.TxtPeriodo.Text = Me.DtaTransacciones.Recordset("Periodo")
     Me.TxtNTransacciones.Text = Me.DtaTransacciones.Recordset("NumeroMovimiento")
     NumeroTransaccion = Me.DtaTransacciones.Recordset("NumeroMovimiento")
     Me.TxtFuente.Text = Me.DtaTransacciones.Recordset("Fuente")
     '//////Sumo los Totales/////////////////////
   
    Debito = 0
    Credito = 0
    TotalDebito = 0
    TotalCredito = 0
      NumFecha1 = Me.TxtFecha.Value
      Fechas = Format(Me.TxtFecha.Value, "yyyy/mm/dd")
      Me.DtaConsulta.RecordSource = "SELECT FechaTransaccion, CodCuentas, NTransaccion, NumeroMovimiento, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, TCambio, Debito, Credito From Transacciones WHERE     (FechaTransaccion = CONVERT(DATETIME, '" & Fechas & "', 102)) AND (NumeroMovimiento = " & NMovimiento & ")"
'
      NMovimiento = Val(Me.TxtNTransacciones)
'      Me.DtaConsulta.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.CodCuentas, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, [Transacciones]![TCambio]*[Transacciones]![Debito] AS MDebito, [Transacciones]![TCambio]*[Transacciones]![Credito] AS MCredito, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito From Transacciones WHERE (((Transacciones.FechaTransaccion)=" & NumFecha1 & ") AND ((Transacciones.NumeroMovimiento)=" & NMovimiento & "))"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
       If Not IsNull(Me.DtaConsulta.Recordset("Debito")) Then
       Debito = Me.DtaConsulta.Recordset("Debito")
       End If
       If Not IsNull(Credito = Me.DtaConsulta.Recordset("Credito")) Then
        Credito = Me.DtaConsulta.Recordset("Credito")
       End If
       TotalDebito = TotalDebito + Debito
       TotalCredito = TotalCredito + Credito
       DtaConsulta.Recordset.MoveNext
      Loop
    Me.TxtCredito.Text = Format(TotalCredito, "##,##0.00")
    Me.TxtDebito.Text = Format(TotalDebito, "##,##0.00")
    Me.TxtDiferencia.Text = Format(TotalDebito - TotalCredito, "##,##0.00")
   

   End If
   
   TxtFecha.Enabled = False
   Me.TxtPeriodo.Enabled = False
   Me.TxtFuente.Enabled = False
   Me.TxtNTransacciones.Enabled = False
   Else
  MsgBox "No existen Transacciones en este Periodo", vbCritical, "Sistema Contable"
  TxtFecha.Enabled = True
   Me.TxtPeriodo.Enabled = True
   Me.TxtFuente.Enabled = True
   Me.TxtNTransacciones.Enabled = True
  Exit Sub
 End If

Else '////////En caso que transaccion tenga un numero en pantalla
     '////////Distinto de Cero////////////////////////
 mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones WHERE (((Transacciones.NombreCuenta)<>'**********CANCELADO*************') AND ((Transacciones.FechaTransaccion) Between '" & Format(NumFecha1, "yyyymmdd") & "' And '" & Format(NumFecha2, "yyyymmdd") & "')) ORDER BY Transacciones.NumeroMovimiento"
 'Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & "))ORDER BY Transacciones.NumeroMovimiento"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
 
 '///////////Busco la Transaccion Anterior////////////
   NumeroAnterior = Me.TxtNTransacciones.Text
   Criterio = "NumeroMovimiento=" & NumeroAnterior & " "
   Me.DtaConsulta.Recordset.Find (Criterio)
   DtaConsulta.Recordset.MovePrevious
 
   If Not DtaConsulta.Recordset.BOF Then
    NumeroTransaccion = DtaConsulta.Recordset("NumeroMovimiento")
   Else '/////en caso que no se encuentre transaccion
    MsgBox "Esta es la primera Transaccion del Periodo", vbInformation, "Sistema Contable"
    Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.NombreCuenta)<>'**********CANCELADO*************') AND ((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion "
     'Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroAnterior & ")) ORDER BY Transacciones.NTransaccion"
     Me.DtaTransacciones.Refresh
   Me.DBGTransacciones.Columns(0).Button = True
     Me.DBGTransacciones.Columns(1).Locked = True
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Button = True
  Me.DBGTransacciones.Columns(6).Locked = True
  Me.DBGTransacciones.Columns(0).Width = 1500
  Me.DBGTransacciones.Columns(2).Width = 1000
  Me.DBGTransacciones.Columns(3).Caption = "Descripcion"
  Me.DBGTransacciones.Columns(4).Width = 1000
  Me.DBGTransacciones.Columns(4).Button = True
  Me.DBGTransacciones.Columns(4).Button = True
  Me.DBGTransacciones.Columns(5).Width = 1000
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Width = 800
  Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
  Me.DBGTransacciones.Columns(7).Locked = True
  Me.DBGTransacciones.Columns(7).NumberFormat = "##,##0.000000"
   Me.DBGTransacciones.Columns(7).Width = 1200
  Me.DBGTransacciones.Columns(8).Width = 1200
  Me.DBGTransacciones.Columns(8).NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns(9).Width = 1200
  Me.DBGTransacciones.Columns(9).NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns(10).Visible = False
  Me.DBGTransacciones.Columns(11).Visible = False
  Me.DBGTransacciones.Columns(12).Visible = False
  Me.DBGTransacciones.Columns(13).Visible = False
  Me.DBGTransacciones.Columns(14).Visible = False
  Me.DBGTransacciones.Columns(15).Visible = False
  Me.DBGTransacciones.Columns(16).Visible = False
 ' Me.DBGTransacciones.Enabled = False
     Exit Sub
   End If
   
    Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.NombreCuenta)<>'**********CANCELADO*************') AND ((Transacciones.FechaTransaccion) Between '" & Format(NumFecha1, "yyyymmdd") & "' And '" & Format(NumFecha2, "yyyymmdd") & "') AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion "
   'Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion"
   Me.DtaTransacciones.Refresh
   If Not DtaTransacciones.Recordset.EOF Then
     Me.TxtFecha.Value = Me.DtaTransacciones.Recordset("FechaTransaccion")
     Me.TxtPeriodo.Text = Me.DtaTransacciones.Recordset("Periodo")
     Me.TxtNTransacciones.Text = Me.DtaTransacciones.Recordset("NumeroMovimiento")
     NumeroTransaccion = Me.DtaTransacciones.Recordset("NumeroMovimiento")
     Me.TxtFuente.Text = Me.DtaTransacciones.Recordset("Fuente")
     
     '/////////////////////////Busco el tipo de moneda del movimiento////////////////
 mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.TipoMoneda,IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between '" & Format(NumFecha1, "yyyymmdd") & "' And '" & Format(NumFecha2, "yyyymmdd") & "' ) AND ((IndiceTransaccion.NumeroMovimiento)= " & NumeroTransaccion & "))"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  If Not IsNull(Me.CmbMoneda.Text = Me.DtaConsulta.Recordset("TipoMoneda")) Then
   Me.CmbMoneda.Text = Me.DtaConsulta.Recordset("TipoMoneda")
  Else
   Me.CmbMoneda.Text = ""
  End If
 End If
     
     
     '//////Sumo los Totales/////////////////////
    
    Debito = 0
    Credito = 0
    TotalDebito = 0
    TotalCredito = 0
      NumFecha1 = Me.TxtFecha.Value
      NMovimiento = Val(Me.TxtNTransacciones)
      Fechas = Format(Me.TxtFecha.Value, "yyyy/mm/dd")
'      NMovimiento = Val(Me.TxtNTransacciones)
      Me.DtaConsulta.RecordSource = "SELECT FechaTransaccion, CodCuentas, NTransaccion, NumeroMovimiento, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, TCambio, Debito, Credito From Transacciones WHERE     (FechaTransaccion = CONVERT(DATETIME, '" & Fechas & "', 102)) AND (NumeroMovimiento = " & NMovimiento & ")"
'
'      Me.DtaConsulta.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.CodCuentas, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, [Transacciones]![TCambio]*[Transacciones]![Debito] AS MDebito, [Transacciones]![TCambio]*[Transacciones]![Credito] AS MCredito, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito From Transacciones WHERE (((Transacciones.FechaTransaccion)=" & NumFecha1 & ") AND ((Transacciones.NumeroMovimiento)=" & NMovimiento & "))"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
      If Not IsNull(Me.DtaConsulta.Recordset("Debito")) Then
       Debito = Me.DtaConsulta.Recordset("Debito")
      End If
      If Not IsNull(Me.DtaConsulta.Recordset("Credito")) Then
       Credito = Me.DtaConsulta.Recordset("Credito")
      End If
       TotalDebito = TotalDebito + Debito
       TotalCredito = TotalCredito + Credito
       DtaConsulta.Recordset.MoveNext
      Loop
    Me.TxtCredito.Text = Format(TotalCredito, "##,##0.00")
    Me.TxtDebito.Text = Format(TotalDebito, "##,##0.00")
    Me.TxtDiferencia.Text = Format(TotalDebito - TotalCredito, "##,##0.00")
   
   Else '////En caso que no encuentre ninguna trasaccion
    
     MsgBox "Esta es la primera Transaccion del Periodo", vbInformation, "Sistema Contable"
     Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroAnterior & ")) ORDER BY Transacciones.NTransaccion"
     Me.DtaTransacciones.Refresh
     
 
   End If
 End If '/////fIN DEL IF CONSULTA////
   TxtFecha.Enabled = False
   Me.TxtPeriodo.Enabled = False
   Me.TxtFuente.Enabled = False
   Me.TxtNTransacciones.Enabled = False
   
  






End If
If Not CodigoUsuario = 0 Then
  Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Transacciones'))"
Me.DtaNacceso.Refresh
If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdGrabar.Enabled = False
   Me.TxtFecha.Enabled = False
   Me.Frame1.Enabled = False
   Me.DBGTransacciones.Enabled = False
Else
  Me.DBGTransacciones.Enabled = True
End If
Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Transacciones'))"
Me.DtaNacceso.Refresh
If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdBorrar.Enabled = False
   Me.SmartButton1.Enabled = False
End If
End If
Salir = True
  Me.DBGTransacciones.Columns(0).Button = True
    Me.DBGTransacciones.Columns(1).Locked = True
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Button = True
  Me.DBGTransacciones.Columns(6).Locked = True
  Me.DBGTransacciones.Columns(0).Width = 1500
  Me.DBGTransacciones.Columns(2).Width = 1000
  Me.DBGTransacciones.Columns(3).Caption = "Descripcion"
  Me.DBGTransacciones.Columns(4).Width = 1000
  Me.DBGTransacciones.Columns(4).Button = True
  Me.DBGTransacciones.Columns(5).Width = 1000
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Width = 800
  Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
  Me.DBGTransacciones.Columns(7).Locked = True
  Me.DBGTransacciones.Columns(7).NumberFormat = "##,##0.000000"
   Me.DBGTransacciones.Columns(7).Width = 1200
  Me.DBGTransacciones.Columns(8).Width = 1200
  Me.DBGTransacciones.Columns(8).NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns(9).Width = 1200
  Me.DBGTransacciones.Columns(9).NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns(10).Visible = False
  Me.DBGTransacciones.Columns(11).Visible = False
  Me.DBGTransacciones.Columns(12).Visible = False
  Me.DBGTransacciones.Columns(13).Visible = False
  Me.DBGTransacciones.Columns(14).Visible = False
  Me.DBGTransacciones.Columns(15).Visible = False
  Me.DBGTransacciones.Columns(16).Visible = False
  'Me.DBGTransacciones.Enabled = False
  
  Exit Sub
TipoErrs:
 MsgBox err.Description
  
  
  
End Sub

Private Sub CmdBorrar_Click()
Dim Periodo As Double

On Error GoTo TipoErrs
  Dim Respuesta, Rsp
  Salir = True
  
  Primero = True
  
  Periodo = NumeroPeriodo
  
  Set Rsp = DtaTransacciones.Recordset
  Respuesta = MsgBox("Esta seguro de Borrar la transaccion?", vbYesNo, "Transaccion No.: " & Me.TxtNTransacciones.Text)
   If Respuesta = 6 Then
   '//////Grabo las descripcion en los indices//////////////////////
   Me.DBGTransacciones.Enabled = True
   mes = Month(Me.TxtFecha.Value)
   Año = Year(Me.TxtFecha.Value)
   FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
   FechaFin = DateSerial(Año, mes + 1, 1 - 1)
   NumFecha1 = FechaIni
   NumFecha2 = FechaFin
 
   Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & " ) AND ((IndiceTransaccion.NumeroMovimiento)= " & NumeroTransaccion & "))"
   Me.DtaConsulta.Refresh
         
       If Not DtaConsulta.Recordset.EOF Then
        
          'Me.'DtaConsulta.Recordset.Edit
          Me.DtaConsulta.Recordset("DescripcionMovimiento") = "*****CANCELADO*****"
          Me.DtaConsulta.Recordset.Update
        
       End If
   
   
   Me.DtaConsulta.RecordSource = "SELECT Transacciones.* From Transacciones WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & " ) AND ((Transacciones.NumeroMovimiento)= " & NumeroTransaccion & ") )"
   Me.DtaConsulta.Refresh
    Do While Not Me.DtaConsulta.Recordset.EOF
'     'Me.DtaTransacciones.Recordset.Edit
     DtaConsulta.Recordset("NombreCuenta") = "**********CANCELADO*************"
     DtaConsulta.Recordset("DescripcionMovimiento") = "**********CANCELADO*************"
     DtaConsulta.Recordset("Beneficiario") = "**********CANCELADO*************"
     DtaConsulta.Recordset("Debito") = 0
     DtaConsulta.Recordset("Credito") = 0

     Me.DtaConsulta.Recordset.Update
     Me.DtaConsulta.Recordset.MoveNext
       
     Me.CmbMoneda.Enabled = False
    Loop
    
'    Me.TxtFecha.Value = Format(FechaSistema, "dd/mm/yyyy")
    Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones Where (((Transacciones.NumeroMovimiento) = -1))"
    Me.DtaTransacciones.Refresh
  Me.DBGTransacciones.Columns(0).Button = True
    Me.DBGTransacciones.Columns(1).Locked = True
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Button = True
  Me.DBGTransacciones.Columns(6).Locked = True
  Me.DBGTransacciones.Columns(0).Width = 1500
  Me.DBGTransacciones.Columns(2).Width = 1000
  Me.DBGTransacciones.Columns(3).Caption = "Descripcion"
  Me.DBGTransacciones.Columns(4).Width = 1000
  Me.DBGTransacciones.Columns(4).Button = True
  Me.DBGTransacciones.Columns(5).Width = 1000
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Width = 800
  Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
  Me.DBGTransacciones.Columns(7).Locked = True
  Me.DBGTransacciones.Columns(7).NumberFormat = "##,##0.000000"
   Me.DBGTransacciones.Columns(7).Width = 1200
  Me.DBGTransacciones.Columns(8).Width = 1200
  Me.DBGTransacciones.Columns(8).NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns(9).Width = 1200
  Me.DBGTransacciones.Columns(9).NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns(10).Visible = False
  Me.DBGTransacciones.Columns(11).Visible = False
  Me.DBGTransacciones.Columns(12).Visible = False
  Me.DBGTransacciones.Columns(13).Visible = False
  Me.DBGTransacciones.Columns(14).Visible = False
  Me.DBGTransacciones.Columns(15).Visible = False
  'Me.DBGTransacciones.Columns(16).Visible = False
   ' Me.DBGTransacciones.Enabled = False
 
 
 If Not CodigoUsuario = 0 Then
Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Transacciones'))"
Me.DtaNacceso.Refresh
If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdGrabar.Enabled = False
   Me.TxtFecha.Enabled = False
   Me.Frame1.Enabled = False
   Me.DBGTransacciones.Enabled = False
   Me.CmdBuscarEmpleado.Enabled = False
      Me.TxtFecha.Enabled = False
Else
  Me.DBGTransacciones.Enabled = True
End If
Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Transacciones'))"
Me.DtaNacceso.Refresh
If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdBorrar.Enabled = False
   Me.SmartButton1.Enabled = False
End If
End If
  
    TotalCredito = 0
    TotalDebito = 0
    Debito = 0
    Credito = 0
    TotalDiferencia = 0
    Diferencia = 0



    TxtFecha.Enabled = True
    Me.TxtPeriodo.Enabled = True
    Me.TxtFuente.Enabled = True
    Me.TxtNTransacciones.Enabled = True

    Me.TxtDebito.Text = "0.00"
    Me.TxtCredito.Text = "0.00"
    Me.TxtDiferencia.Text = "0.00"
    Me.TxtFuente.Text = ""
    Me.TxtNTransacciones.Text = "0"
    TxtFecha.Enabled = True
    Me.TxtPeriodo.Enabled = True
    Me.TxtFuente.Enabled = True
    Me.TxtNTransacciones.Enabled = True
  End If
 Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub CmdBuscarEmpleado_Click()
 QueProducto = "NTransacciones"
 FrmConsulta.Show 1
 
End Sub

Private Sub CmdCancelar_Click()
Me.DBGTransacciones.Enabled = True
  TDBGridFechas.Visible = False
  
End Sub

Private Sub CmdGrabar_Click()
On Error GoTo TipoErrs

Dim Sql As String

Me.TDBGridFechas.Visible = False

 If Not Val(Me.TxtDiferencia.Text) = 0 Then
   MsgBox "El Movimiento esta Desbalanceado", vbCritical, "Sistema Contable"
   Exit Sub
  End If
Me.CmdNuevo.Enabled = True
Me.CmbMoneda.Enabled = True
Me.CmdBuscarEmpleado.Enabled = True
'//////Grabo las descripcion en los indices//////////////////////
 Me.DBGTransacciones.Enabled = True
 mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 If Me.CmbMoneda.Text = "" Then
   Me.CmbMoneda.Text = "Córdobas"
 End If
 

 
Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.TipoMoneda,IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente From IndiceTransaccion WHERE IndiceTransaccion.FechaTransaccion>='" & Format(NumFecha1, "yyyymmdd") & "' And IndiceTransaccion.FechaTransaccion<='" & Format(NumFecha2, "yyyymmdd") & "' AND IndiceTransaccion.NumeroMovimiento= " & NumeroTransaccion
 'la cambie por el mismo problema del between
 'Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.TipoMoneda,IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & " ) AND ((IndiceTransaccion.NumeroMovimiento)= " & NumeroTransaccion & "))"
 Me.DtaConsulta.Refresh
 If Not DtaTransacciones.Recordset.EOF Then
 Me.DtaTransacciones.Recordset.MoveFirst
 End If
       
       If Not DtaConsulta.Recordset.EOF Then
  
          'Me.'DtaConsulta.Recordset.Edit
          Me.DtaConsulta.Recordset("TipoMoneda") = Me.CmbMoneda.Text
         If Not Me.DBGTransacciones.Columns(3).Text = "" Then
          Me.DtaConsulta.Recordset("DescripcionMovimiento") = Me.DBGTransacciones.Columns(3).Text
         End If
          Me.DtaConsulta.Recordset.Update

       End If
       
   Primero = True
Me.TxtDebito.Text = "0.00"
Me.TxtCredito.Text = "0.00"
Me.TxtDiferencia.Text = "0.00"
Me.TxtFuente.Text = ""
Me.TxtNTransacciones.Text = "0"
TxtFecha.Enabled = True
Me.TxtPeriodo.Enabled = True
Me.TxtFuente.Enabled = True
Me.TxtNTransacciones.Enabled = True
        
       
If Not CodigoUsuario = 0 Then
Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Transacciones'))"
Me.DtaNacceso.Refresh
If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdGrabar.Enabled = False
   Me.TxtFecha.Enabled = False
   Me.Frame1.Enabled = False
   Me.DBGTransacciones.Enabled = False
      Me.TxtFecha.Enabled = False
Else
  Me.DBGTransacciones.Enabled = True
End If
Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Transacciones'))"
Me.DtaNacceso.Refresh
If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdBorrar.Enabled = False
   Me.SmartButton1.Enabled = False
End If

Sql = "SELECT     Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, " & _
       "Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, " & _
       "Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, " & _
       "Transacciones.NumeroMovimiento, Periodos.Periodo, Transacciones.FechaDescuento, Transacciones.DescuentoDisponible, " & _
       "Transacciones.FechaVence,Transacciones.CodCuentaProveedor,Transacciones.TipoFactura,Transacciones.NTransaccion " & _
       "FROM         Periodos INNER JOIN " & _
       "Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo " & _
       "Where (Transacciones.NumeroMovimiento = -1) " & _
       "ORDER BY Transacciones.NTransaccion "

'Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones Where (((Transacciones.NumeroMovimiento) = -1))"
Me.DtaTransacciones.RecordSource = Sql
Me.DtaTransacciones.Refresh
End If


Salir = True
Sql = "SELECT     Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, " & _
       "Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, " & _
       "Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, " & _
       "Transacciones.NumeroMovimiento, Periodos.Periodo, Transacciones.FechaDescuento, Transacciones.DescuentoDisponible, " & _
       "Transacciones.FechaVence,Transacciones.CodCuentaProveedor,Transacciones.TipoFactura,Transacciones.NTransaccion " & _
       "FROM         Periodos INNER JOIN " & _
       "Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo " & _
       "Where (Transacciones.NumeroMovimiento = -1) " & _
       "ORDER BY Transacciones.NTransaccion "
       
'Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones Where (((Transacciones.NumeroMovimiento) = -1))"
Me.DtaTransacciones.RecordSource = Sql
Me.DtaTransacciones.Refresh
Me.CmbMoneda.Text = "Córdobas"

  Me.DBGTransacciones.Columns(0).Button = True
  Me.DBGTransacciones.Columns(1).Locked = True
  Me.DBGTransacciones.Columns(1).Locked = True
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Button = True
  Me.DBGTransacciones.Columns(6).Locked = True
  Me.DBGTransacciones.Columns(0).Width = 1500
  Me.DBGTransacciones.Columns(2).Width = 1000
  Me.DBGTransacciones.Columns(3).Caption = "Descripcion"
  Me.DBGTransacciones.Columns(4).Width = 1000
  Me.DBGTransacciones.Columns(4).Button = True
  Me.DBGTransacciones.Columns(5).Width = 1000
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Width = 800
  Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
  Me.DBGTransacciones.Columns(7).Locked = True
  Me.DBGTransacciones.Columns(7).NumberFormat = "##,##0.000000"
  Me.DBGTransacciones.Columns(7).Width = 1200
  Me.DBGTransacciones.Columns(8).Width = 1200
  Me.DBGTransacciones.Columns(8).NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns(9).Width = 1200
  Me.DBGTransacciones.Columns(9).NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns(10).Visible = False
  Me.DBGTransacciones.Columns(11).Visible = False
  Me.DBGTransacciones.Columns(12).Visible = False
  Me.DBGTransacciones.Columns(13).Visible = False
  Me.DBGTransacciones.Columns(14).Visible = False
  Me.DBGTransacciones.Columns(15).Visible = False
  Me.DBGTransacciones.Columns(16).Visible = False
  Me.DBGTransacciones.Columns(17).Visible = False
  Me.DBGTransacciones.Columns(18).Visible = False
  Me.DBGTransacciones.Columns(19).Visible = False
  FrmTransacciones.DBGTransacciones.Columns(20).Visible = False
  FrmTransacciones.DBGTransacciones.Columns(21).Visible = False
  FrmTransacciones.DBGTransacciones.Columns(22).Visible = False
  Me.DBGTransacciones.Columns(7).Locked = True 'columna tasa de cambio
  
  
TotalCredito = 0
TotalDebito = 0
Debito = 0
Credito = 0
TotalDiferencia = 0
Diferencia = 0


'////////////////ACTUALIZO LAS FUENTES///////////////////////////////////////////////////7
Me.AdoFuente.RecordSource = "SELECT DISTINCT Fuente From IndiceTransaccion WHERE (Fuente <> N'  ')"
Me.AdoFuente.Refresh

 
 Exit Sub
TipoErrs:
 MsgBox err.Description
End Sub

Private Sub CmdImprimir_Click()
Dim Sql As String
Dim NumeroMovimientos As Double
 If Me.TxtNTransacciones.Text = "" Or Me.TxtNTransacciones.Text = "0" Then Exit Sub
  
  NumeroMovimientos = Me.TxtNTransacciones.Text
  
 ArepTransacciones.LblFecha = Format(Now, "dd/mm/yyyy")
    ArepTransacciones.LblRangoFecha = Format(TxtFecha, "long date")
    
    ArepTransacciones.DataControl1.ConnectionString = ConexionReporte
        ArepTransacciones.LblNombre.Caption = "Comprobantes de Diario"
        Sql = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta AS DescripcionCuentas, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, " & _
            "Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, " & _
            "Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, " & _
            "Transacciones.NumeroMovimiento , Periodos.Periodo " & _
            "FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo " & _
            "WHERE  (Transacciones.FechaTransaccion BETWEEN '" & Format(Me.TxtFecha.Value, "yyyymmdd") & "' And '" & Format(Me.TxtFecha.Value, "yyyymmdd") & "') AND (Transacciones.NumeroMovimiento = " & NumeroMovimientos & ") " & _
            "ORDER BY Transacciones.NTransaccion"
            
        ArepTransacciones.DataControl1.Source = Sql
        

        
        If UCase(Me.TxtFuente.Text) = UCase("Cheque") Then
            ArepTransacciones.LblNombre.Caption = "Comprobantes de Pago"
        Else
            ArepTransacciones.LblNombre.Caption = "Comprobante de Diario"
        End If


    
    ArepTransacciones.LblEmpresa = MDIPrimero.AdoConfiguracion.Recordset("NombreEmpresa")
    ArepTransacciones.LblEmpresa1 = MDIPrimero.AdoConfiguracion.Recordset("Direccion")
    ArepTransacciones.LblEmpresa2 = "RUC: " & MDIPrimero.AdoConfiguracion.Recordset("NumeroRuc")
'    ArepTransacciones.Logo.Picture = LoadPicture(RutaLogo)
    ArepTransacciones.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
    ArepTransacciones.Moneda = Me.CmbMoneda.Text
'    ArepTransacciones.Show 1

Dim rpt As Object
Dim fPreview As New FrmPreview

     Set rpt = New ArepTransacciones
     rpt.DataControl1.ConnectionString = ConexionReporte
     rpt.DataControl1.Source = Sql
     fPreview.RunReport rpt


     fPreview.Show 1


End Sub

Private Sub CmdNuevo_Click()
On Error GoTo TipoErrs
  If Not Val(Me.TxtDiferencia.Text) = 0 Then
   MsgBox "El Movimiento esta Desbalanceado", vbCritical, "Sistema Contable"
   Exit Sub
  End If
 Me.CmdBuscarEmpleado.Enabled = False
 Me.CmbMoneda.Enabled = True
 '//////Grabo las descripcion en los indices//////////////////////
 Me.DBGTransacciones.Enabled = True
 mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & " ) AND ((IndiceTransaccion.NumeroMovimiento)= " & NumeroTransaccion & "))"
 'problema con el between
    Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente From IndiceTransaccion WHERE IndiceTransaccion.FechaTransaccion>='" & Format(NumFecha1, "yyyymmdd") & "' And IndiceTransaccion.FechaTransaccion<='" & Format(NumFecha2, "yyyymmdd") & "' AND IndiceTransaccion.NumeroMovimiento= " & NumeroTransaccion
    Me.DtaConsulta.Refresh
 If Not DtaTransacciones.Recordset.EOF Then
 Me.DtaTransacciones.Recordset.MoveFirst
 End If
       
       If Not DtaConsulta.Recordset.EOF Then
        If Not Me.DBGTransacciones.Columns(3).Text = "" Then
          'Me.'DtaConsulta.Recordset.Edit
          Me.DtaConsulta.Recordset("DescripcionMovimiento") = Me.DBGTransacciones.Columns(3).Text
          Me.DtaConsulta.Recordset.Update
        End If
       End If
       
  TxtFecha.Enabled = True
Me.TxtPeriodo.Enabled = True
Me.TxtFuente.Enabled = True
Me.TxtNTransacciones.Enabled = True
Primero = True
Me.TxtDebito.Text = "0.00"
Me.TxtCredito.Text = "0.00"
Me.TxtDiferencia.Text = "0.00"
Me.TxtFuente.Text = ""
Me.TxtNTransacciones.Text = "0"
TxtFecha.Enabled = True
Me.TxtPeriodo.Enabled = True
Me.TxtFuente.Enabled = True
Me.TxtNTransacciones.Enabled = True
 
If Not CodigoUsuario = 0 Then
 Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Transacciones'))"
Me.DtaNacceso.Refresh
If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdGrabar.Enabled = False
   Me.TxtFecha.Enabled = False
   Me.Frame1.Enabled = False
   Me.DBGTransacciones.Enabled = False
   Me.TxtFecha.Enabled = False
Else
  Me.DBGTransacciones.Enabled = True
End If
Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Transacciones'))"
Me.DtaNacceso.Refresh
If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdBorrar.Enabled = False
   Me.SmartButton1.Enabled = False
End If
 End If
Salir = True

Sql = "SELECT     Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, " & _
       "Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, " & _
       "Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, " & _
       "Transacciones.NumeroMovimiento, Periodos.Periodo, Transacciones.FechaDescuento, Transacciones.DescuentoDisponible, " & _
       "Transacciones.FechaVence,Transacciones.CodCuentaProveedor,Transacciones.TipoFactura,Transacciones.NTransaccion " & _
       "FROM         Periodos INNER JOIN " & _
       "Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo " & _
       "Where (Transacciones.NumeroMovimiento = -1) " & _
       "ORDER BY Transacciones.NTransaccion "
       
Me.DtaTransacciones.RecordSource = Sql
Me.DtaTransacciones.Refresh
Me.CmbMoneda.Text = "Córdobas"

  Me.DBGTransacciones.Columns(0).Button = True
  Me.DBGTransacciones.Columns(1).Locked = True
  Me.DBGTransacciones.Columns(1).Locked = True
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Button = True
  Me.DBGTransacciones.Columns(6).Locked = True
  Me.DBGTransacciones.Columns(0).Width = 1500
  Me.DBGTransacciones.Columns(2).Width = 1000
  Me.DBGTransacciones.Columns(3).Caption = "Descripcion"
  Me.DBGTransacciones.Columns(4).Width = 1000
  Me.DBGTransacciones.Columns(4).Button = True
  Me.DBGTransacciones.Columns(5).Width = 1000
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Width = 800
  Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
  Me.DBGTransacciones.Columns(7).Locked = True
  Me.DBGTransacciones.Columns(7).NumberFormat = "##,##0.000000"
  Me.DBGTransacciones.Columns(7).Width = 1200
  Me.DBGTransacciones.Columns(8).Width = 1200
  Me.DBGTransacciones.Columns(8).NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns(9).Width = 1200
  Me.DBGTransacciones.Columns(9).NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns(10).Visible = False
  Me.DBGTransacciones.Columns(11).Visible = False
  Me.DBGTransacciones.Columns(12).Visible = False
  Me.DBGTransacciones.Columns(13).Visible = False
  Me.DBGTransacciones.Columns(14).Visible = False
  Me.DBGTransacciones.Columns(15).Visible = False
  Me.DBGTransacciones.Columns(16).Visible = False
  Me.DBGTransacciones.Columns(17).Visible = False
  Me.DBGTransacciones.Columns(18).Visible = False
  Me.DBGTransacciones.Columns(19).Visible = False
  FrmTransacciones.DBGTransacciones.Columns(20).Visible = False
  FrmTransacciones.DBGTransacciones.Columns(21).Visible = False
  FrmTransacciones.DBGTransacciones.Columns(22).Visible = False
  Me.DBGTransacciones.Columns(7).Locked = True 'columna tasa de cambio

TotalCredito = 0
TotalDebito = 0
Debito = 0
Credito = 0
TotalDiferencia = 0
Diferencia = 0




 
 Exit Sub
TipoErrs:
 MsgBox err.Description
End Sub

Private Sub CmdSalir_Click()
On Error GoTo TipoErrs
'//////Grabo las descripcion en los indices//////////////////////
 Me.DBGTransacciones.Enabled = True
 mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente From IndiceTransaccion WHERE IndiceTransaccion.FechaTransaccion>='" & Format(NumFecha1, "yyyymmdd") & "' And IndiceTransaccion.FechaTransaccion<='" & Format(NumFecha2, "yyyymmdd") & "' AND IndiceTransaccion.NumeroMovimiento= " & NumeroTransaccion
  'between
'  Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & " ) AND ((IndiceTransaccion.NumeroMovimiento)= " & NumeroTransaccion & "))"
 Me.DtaConsulta.Refresh
 If Not DtaTransacciones.Recordset.EOF Then
 Me.DtaTransacciones.Recordset.MoveFirst
 End If
       If Not DtaConsulta.Recordset.EOF Then
        If Not Me.DBGTransacciones.Columns(3).Text = "" Then
          'Me.'DtaConsulta.Recordset.Edit
          Me.DtaConsulta.Recordset("DescripcionMovimiento") = Me.DBGTransacciones.Columns(3).Text
          Me.DtaConsulta.Recordset.Update
        End If
       End If
Unload Me
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub DBCombo1_Click(Area As Integer)

End Sub

Private Sub DBCombo1_GotFocus()

End Sub




Private Sub CmdSiguiente_Click()
Dim Fechas1 As String, Fechas2 As String, NumeroMovimiento As Integer
'Dim SQL As String
'On Error GoTo TipoErrs
 If Not Val(Me.TxtDiferencia.Text) = 0 Then
   MsgBox "El Movimiento esta Desbalanceado", vbCritical, "Sistema Contable"
   Exit Sub
  End If
 Primero = True
  Me.CmbMoneda.Enabled = False
  
  '//////Grabo las descripcion en los indices//////////////////////
 Me.DBGTransacciones.Enabled = True
 mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 Fechas1 = Format(FechaIni, "yyyy/mm/dd")
 Fechas2 = Format(FechaFin, "yyyy/mm/dd")
Me.DtaConsulta.RecordSource = "SELECT TipoMoneda, FechaTransaccion, NumeroMovimiento, DescripcionMovimiento, Fuente From IndiceTransaccion WHERE  (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas2 & "', 102)) AND  (NumeroMovimiento = " & NumeroTransaccion & ")"
' Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.TipoMoneda,IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & " ) AND ((IndiceTransaccion.NumeroMovimiento)= " & NumeroTransaccion & "))"
 Me.DtaConsulta.Refresh
 If Not DtaTransacciones.Recordset.EOF Then
 Me.DtaTransacciones.Recordset.MoveFirst
 End If
       
       If Not DtaConsulta.Recordset.EOF Then
         If Not IsNull(Me.CmbMoneda.Text = Me.DtaConsulta.Recordset("TipoMoneda")) Then
          'Me.CmbMoneda.Text = Me.DtaConsulta.Recordset("TipoMoneda")
         End If
  
        If Not Me.DBGTransacciones.Columns(3).Text = "" Then
          'Me.'DtaConsulta.Recordset.Edit
'          Me.DtaConsulta.Recordset("TipoMoneda") = Me.CmbMoneda.Text
'          Me.DtaConsulta.Recordset("DescripcionMovimiento") = Me.DBGTransacciones.Columns(3).Text
'          Me.DtaConsulta.Recordset.Update
        Else
          'Me.'DtaConsulta.Recordset.Edit
          'Me.DtaConsulta.Recordset("TipoMoneda") = Me.CmbMoneda.Text
'          Me.DtaConsulta.Recordset.Update
        End If
       End If
       
 TotalDiferencia = 0
 TotalCredito = 0
 TotalDebito = 0
 Debito = 0
 Credito = 0
 Diferencia = 0
 
If Me.TxtNTransacciones = 0 Then
 mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones WHERE (((Transacciones.NombreCuenta)<>'**********CANCELADO*************') AND ((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ")) ORDER BY Transacciones.NumeroMovimiento"
 'Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & "))ORDER BY Transacciones.NumeroMovimiento"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
   '/////////Me muevo al ultimo registro/////////
   Me.DtaConsulta.Recordset.MoveLast
   NumeroTransaccion = DtaConsulta.Recordset("NumeroMovimiento")
   Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.NombreCuenta)<>'**********CANCELADO*************') AND ((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion "
   'Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion"
   Me.DtaTransacciones.Refresh
   
   If Not DtaTransacciones.Recordset.EOF Then
     Me.TxtFecha.Value = Me.DtaTransacciones.Recordset("FechaTransaccion")
     Me.TxtPeriodo.Text = Me.DtaTransacciones.Recordset("Periodo")
     Me.TxtNTransacciones.Text = Me.DtaTransacciones.Recordset("NumeroMovimiento")
     NumeroTransaccion = Me.DtaTransacciones.Recordset("NumeroMovimiento")
     Me.TxtFuente.Text = Me.DtaTransacciones.Recordset("Fuente")
         '/////////////////////////Busco el tipo de moneda del movimiento////////////////
 mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.TipoMoneda,IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & " ) AND ((IndiceTransaccion.NumeroMovimiento)= " & NumeroTransaccion & "))"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  If Not IsNull(Me.DtaConsulta.Recordset("TipoMoneda")) Then
   Me.CmbMoneda.Text = Me.DtaConsulta.Recordset("TipoMoneda")
  Else
   Me.CmbMoneda.Text = ""
  End If
 End If
     
     '//////Sumo los Totales/////////////////////
    Debito = 0
    Credito = 0
    TotalDebito = 0
    TotalCredito = 0
      NumFecha1 = Me.TxtFecha.Value
      NMovimiento = Val(Me.TxtNTransacciones)
      Fechas = Format(Me.TxtFecha.Value, "yyyy/mm/dd")
       Me.DtaConsulta.RecordSource = "SELECT FechaTransaccion, CodCuentas, NTransaccion, NumeroMovimiento, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, TCambio, Debito, Credito From Transacciones WHERE     (FechaTransaccion = CONVERT(DATETIME, '" & Fechas & "', 102)) AND (NumeroMovimiento = " & NMovimiento & ")"
'
'      Me.DtaConsulta.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.CodCuentas, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, [Transacciones]![TCambio]*[Transacciones]![Debito] AS MDebito, [Transacciones]![TCambio]*[Transacciones]![Credito] AS MCredito, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito From Transacciones WHERE (((Transacciones.FechaTransaccion)=" & NumFecha1 & ") AND ((Transacciones.NumeroMovimiento)=" & NMovimiento & "))"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
       If Not IsNull(Me.DtaConsulta.Recordset("Debito")) Then
       Debito = Me.DtaConsulta.Recordset("Debito")
       End If
       If Not IsNull(Credito = Me.DtaConsulta.Recordset("Credito")) Then
        Credito = Me.DtaConsulta.Recordset("Credito")
       End If
       TotalDebito = TotalDebito + Debito
       TotalCredito = TotalCredito + Credito
       DtaConsulta.Recordset.MoveNext
      Loop
    Me.TxtCredito.Text = Format(TotalCredito, "##,##0.00")
    Me.TxtDebito.Text = Format(TotalDebito, "##,##0.00")
    Me.TxtDiferencia.Text = Format(TotalDebito - TotalCredito, "##,##0.00")
   
   End If
   
   TxtFecha.Enabled = False
   Me.TxtPeriodo.Enabled = False
   Me.TxtFuente.Enabled = False
   Me.TxtNTransacciones.Enabled = False
   Else
  MsgBox "No existen Transacciones en este Periodo", vbCritical, "Sistema Contable"
  TxtFecha.Enabled = True
   Me.TxtPeriodo.Enabled = True
   Me.TxtFuente.Enabled = True
   Me.TxtNTransacciones.Enabled = True
  Exit Sub
 End If

Else '////////En caso que transaccion tenga un numero en pantalla
     '////////Distinto de Cero////////////////////////
 mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
  Fechas1 = Format(FechaIni, "yyyy/mm/dd")
  Fechas2 = Format(FechaFin, "yyyy/mm/dd")
  
 
'  SQL = "SELECT  Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, " & vbLf
'  SQL = SQL & "Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito," & vbLf
'  SQL = SQL & "Transacciones.FechaTransaccion , Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, " & vbLf
'  SQL = SQL & "Transacciones.NumeroMovimiento , Periodos.Periodo" & vbLf
'  SQL = SQL & "FROM Periodos INNER JOIN" & vbLf
'  SQL = SQL & "Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo" & vbLf
'  SQL = SQL & "WHERE     (Transacciones.NombreCuenta <> '**********CANCELADO*************')" & vbLf
'  SQL = SQL & "AND (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME,'" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas2 & "', 102))" & vbLf
'  SQL = SQL & "ORDER BY Transacciones.NumeroMovimiento"
'
  
'  SQL = "SELECT FechaTransaccion, NumeroMovimiento, DescripcionMovimiento, Nperiodo, Fuente, TipoMoneda From IndiceTransaccion WHERE     (DescripcionMovimiento <> N'*****CANCELADO*****') AND (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas2 & "', 102)) ORDER BY NumeroMovimiento"
  Sql = "SELECT DISTINCT NumeroMovimiento, NPeriodo, FechaTransaccion From Transacciones WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas2 & "', 102)) ORDER BY NumeroMovimiento"
  Me.DtaConsulta.RecordSource = Sql
  
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
'  NumeroMovimiento= me.DtaConsulta.Recordset.
 '///////////Busco la Transaccion Siguiente////////////
   NumeroAnterior = Me.TxtNTransacciones.Text
    Fechas1 = Format(FechaIni, "yyyy/mm/dd")
    Fechas2 = Format(FechaFin, "yyyy/mm/dd")
 
   Criterio = "NumeroMovimiento=" & NumeroAnterior & " "
   Me.DtaConsulta.Recordset.MoveFirst
   Me.DtaConsulta.Recordset.Find (Criterio)
   DtaConsulta.Recordset.MoveNext
   
 If Not DtaConsulta.Recordset.EOF Then
   NumeroTransaccion = DtaConsulta.Recordset("NumeroMovimiento")
 Else '//En caso que no se encuentre ninguna transaccion
  MsgBox "Esta es la ultima Transaccion del Periodo", vbInformation, "Sistema Contable"
      Fechas1 = Format(FechaIni, "yyyy/mm/dd")
    Fechas2 = Format(FechaFin, "yyyy/mm/dd")
   Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.NombreCuenta)<>'**********CANCELADO*************') AND (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas2 & "', 102)) AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion "
'    Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.NombreCuenta)<>'**********CANCELADO*************') AND ((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion "
    ' Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,  Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroAnterior & ")) ORDER BY Transacciones.NTransaccion"
     Me.DtaTransacciones.Refresh
   Me.DBGTransacciones.Columns(0).Button = True
     Me.DBGTransacciones.Columns(1).Locked = True
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Button = True
  Me.DBGTransacciones.Columns(6).Locked = True
  Me.DBGTransacciones.Columns(0).Width = 1500
  Me.DBGTransacciones.Columns(2).Width = 1000
  Me.DBGTransacciones.Columns(3).Caption = "Descripcion"
  Me.DBGTransacciones.Columns(4).Width = 1000
  Me.DBGTransacciones.Columns(4).Button = True
  Me.DBGTransacciones.Columns(5).Width = 1000
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Width = 800
  Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
  Me.DBGTransacciones.Columns(7).Locked = True
  Me.DBGTransacciones.Columns(7).NumberFormat = "##,##0.000000"
  Me.DBGTransacciones.Columns(7).Width = 1200
  Me.DBGTransacciones.Columns(8).Width = 1200
  Me.DBGTransacciones.Columns(8).NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns(9).Width = 1200
  Me.DBGTransacciones.Columns(9).NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns(10).Visible = False
  Me.DBGTransacciones.Columns(11).Visible = False
  Me.DBGTransacciones.Columns(12).Visible = False
  Me.DBGTransacciones.Columns(13).Visible = False
  Me.DBGTransacciones.Columns(14).Visible = False
  Me.DBGTransacciones.Columns(15).Visible = False
  Me.DBGTransacciones.Columns(16).Visible = False
  'Me.DBGTransacciones.Enabled = False
     Exit Sub
 End If
    NumFecha1 = FechaIni
    NumFecha2 = FechaFin
    Fechas1 = Format(FechaIni, "yyyy/mm/dd")
    Fechas2 = Format(FechaFin, "yyyy/mm/dd")
   Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.NombreCuenta)<>'**********CANCELADO*************') AND (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas2 & "', 102)) AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion "
   'Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,  Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion"
   Me.DtaTransacciones.Refresh
   If Not DtaTransacciones.Recordset.EOF Then
     Me.TxtFecha.Value = Me.DtaTransacciones.Recordset("FechaTransaccion")
     Me.TxtPeriodo.Text = Me.DtaTransacciones.Recordset("Periodo")
     Me.TxtNTransacciones.Text = Me.DtaTransacciones.Recordset("NumeroMovimiento")
      NumeroTransaccion = Me.DtaTransacciones.Recordset("NumeroMovimiento")
     Me.TxtFuente.Text = Me.DtaTransacciones.Recordset("Fuente")
     
     '/////////////////////////Busco el tipo de moneda del movimiento////////////////
 mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
    Fechas1 = Format(FechaIni, "yyyy/mm/dd")
    Fechas2 = Format(FechaFin, "yyyy/mm/dd")
 
 Me.DtaConsulta.RecordSource = "SELECT TipoMoneda, FechaTransaccion, NumeroMovimiento, DescripcionMovimiento, Fuente From IndiceTransaccion WHERE     (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas2 & "', 102)) AND (NumeroMovimiento = " & NumeroTransaccion & ")"
 'Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.TipoMoneda,IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & " ) AND ((IndiceTransaccion.NumeroMovimiento)= " & NumeroTransaccion & "))"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  If Not IsNull(Me.DtaConsulta.Recordset("TipoMoneda")) Then
   Me.CmbMoneda.Text = Me.DtaConsulta.Recordset("TipoMoneda")
  Else
   Me.CmbMoneda.Text = ""
  End If
 End If
     
     
     '//////Sumo los Totales/////////////////////
   
    Debito = 0
    Credito = 0
    TotalDebito = 0
    TotalCredito = 0
      NumFecha1 = Me.TxtFecha.Value
      NMovimiento = Val(Me.TxtNTransacciones)
      Fechas = Format(Me.TxtFecha.Value, "yyyy/mm/dd")
      Me.DtaConsulta.RecordSource = "SELECT FechaTransaccion, CodCuentas, NTransaccion, NumeroMovimiento, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, TCambio, Debito, Credito From Transacciones WHERE     (FechaTransaccion = CONVERT(DATETIME, '" & Fechas & "', 102)) AND (NumeroMovimiento = " & NMovimiento & ")"
'
'      Me.DtaConsulta.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.CodCuentas, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, [Transacciones]![TCambio]*[Transacciones]![Debito] AS MDebito, [Transacciones]![TCambio]*[Transacciones]![Credito] AS MCredito, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito From Transacciones WHERE (((Transacciones.FechaTransaccion)=" & NumFecha1 & ") AND ((Transacciones.NumeroMovimiento)=" & NMovimiento & "))"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
      If Not IsNull(Me.DtaConsulta.Recordset("Debito")) Then
       Debito = Me.DtaConsulta.Recordset("Debito")
      End If
      If Not IsNull(Me.DtaConsulta.Recordset("Credito")) Then
       Credito = Me.DtaConsulta.Recordset("Credito")
      End If
       TotalDebito = TotalDebito + Debito
       TotalCredito = TotalCredito + Credito
       DtaConsulta.Recordset.MoveNext
      Loop
    Me.TxtCredito.Text = Format(TotalCredito, "##,##0.00")
    Me.TxtDebito.Text = Format(TotalDebito, "##,##0.00")
    Me.TxtDiferencia.Text = Format(TotalDebito - TotalCredito, "##,##0.00")
   
   Else '////En caso que no encuentre ninguna trasaccion
    
     MsgBox "Esta es la ultima Transaccion del Periodo", vbInformation, "Sistema Contable"
     Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,  Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE Transacciones.FechaTransaccion>='" & Format(NumFecha1, "yyyymmdd") & "' And Transacciones.FechaTransaccion<='" & Format(NumFecha2, "yyyymmdd") & "' AND Transacciones.NumeroMovimiento=" & NumeroAnterior & " ORDER BY Transacciones.NTransaccion"
     Me.DtaTransacciones.Refresh
     
   
   
   
   End If
   
 End If '/////////Fin del if consulta////////
   TxtFecha.Enabled = False
   Me.TxtPeriodo.Enabled = False
   Me.TxtFuente.Enabled = False
   Me.TxtNTransacciones.Enabled = False
   
  






End If
If Not CodigoUsuario = 0 Then
 Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Transacciones'))"
Me.DtaNacceso.Refresh
If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdGrabar.Enabled = False
   Me.TxtFecha.Enabled = False
   Me.Frame1.Enabled = False
   Me.DBGTransacciones.Enabled = False
Else
  Me.DBGTransacciones.Enabled = True
End If
Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Transacciones'))"
Me.DtaNacceso.Refresh
If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdBorrar.Enabled = False
   Me.SmartButton1.Enabled = False
End If
End If


  Salir = True
   Me.DBGTransacciones.Columns(0).Button = True
     Me.DBGTransacciones.Columns(1).Locked = True
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Button = True
  Me.DBGTransacciones.Columns(6).Locked = True
  Me.DBGTransacciones.Columns(0).Width = 1500
  Me.DBGTransacciones.Columns(2).Width = 1000
  Me.DBGTransacciones.Columns(3).Caption = "Descripcion"
  Me.DBGTransacciones.Columns(4).Width = 1000
  Me.DBGTransacciones.Columns(4).Button = True
  Me.DBGTransacciones.Columns(5).Width = 1000
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Width = 800
  Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
  Me.DBGTransacciones.Columns(7).Locked = True
  Me.DBGTransacciones.Columns(7).NumberFormat = "##,##0.000000"
  Me.DBGTransacciones.Columns(7).Width = 1200
  Me.DBGTransacciones.Columns(8).Width = 1200
  Me.DBGTransacciones.Columns(8).NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns(9).Width = 1200
  Me.DBGTransacciones.Columns(9).NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns(10).Visible = False
  Me.DBGTransacciones.Columns(11).Visible = False
  Me.DBGTransacciones.Columns(12).Visible = False
  Me.DBGTransacciones.Columns(13).Visible = False
  Me.DBGTransacciones.Columns(14).Visible = False
  Me.DBGTransacciones.Columns(15).Visible = False
    Me.DBGTransacciones.Columns(16).Visible = False
  'Me.DBGTransacciones.Enabled = False
  
  Exit Sub
TipoErrs:
 MsgBox err.Description
End Sub


Private Sub Combo1_Change()

End Sub

Private Sub Command1_Click()
QueProducto = "CuentaFactura"
FrmConsulta.Show 1
End Sub

Private Sub DBGTransacciones_AfterColEdit(ByVal ColIndex As Integer)
On Error GoTo TipoErrs
Dim Descripcion As String, TipoCunta As String, numero As String, Fecha As Long
Dim MontoTasa As Double, CodigoCuenta As String, TipoCuenta As String
Dim DescripcionMovimiento As String, Sql As String, ChequeNo As String

'Este Procedimiento es solo cuando se ejecuta directamente de Recepcion
QueProducto = "Transacciones"
Me.CmbMoneda.Enabled = False

If ColIndex = 8 Or ColIndex = 9 Then
 If Me.DBGTransacciones.Columns(6).Text = "Debito" Then
    If Not ColIndex = 8 Then
      Me.DBGTransacciones.Columns(ColIndex) = "0.00"
    End If
      'Me.DBGTransacciones.Columns(9).Locked = True
      'Me.DBGTransacciones.Columns(8).Locked = False
  ElseIf Me.DBGTransacciones.Columns(6).Text = "Credito" Then
      If Not ColIndex = 9 Then
        Me.DBGTransacciones.Columns(ColIndex) = "0.00"
      End If
     'Me.DBGTransacciones.Columns(9).Locked = False
     'Me.DBGTransacciones.Columns(8).Locked = True
  End If
 End If



  If ColIndex = 4 Then
  
      
        If Me.ChkVentana.Value = 1 Then
          '//////////////////////////////////////////////////////////
          '////BUSCO EL TIPO DE CUENTA PARA MOSTRAR FECHAS///////////
          '////////////////////////////////////////////////////////
          
           CodigoCuenta = Me.DBGTransacciones.Columns("CodCuentas").Text
           Me.DtaConsulta.RecordSource = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo From Cuentas WHERE (CodCuentas = '" & CodigoCuenta & "')"
           Me.DtaConsulta.Refresh
           If Not Me.DtaConsulta.Recordset.EOF Then
             TipoCuenta = Me.DtaConsulta.Recordset("TipoCuenta")
             If TipoCuenta = "Cuentas x Cobrar" Then
                Set c = DBGTransacciones.Columns(ColIndex)
                With Me.TDBGridFechas
                    .Left = Me.DBGTransacciones.Left + c.Left
                    .top = DBGTransacciones.top + DBGTransacciones.RowTop(DBGTransacciones.Row) + DBGTransacciones.RowHeight
                    .Visible = True
                    .SetFocus
        
                End With
             ElseIf TipoCuenta = "Cuentas x Pagar" Then
                Set c = DBGTransacciones.Columns(ColIndex)
                With Me.TDBGridFechas
                    .Left = Me.DBGTransacciones.Left + c.Left
                    .top = DBGTransacciones.top + DBGTransacciones.RowTop(DBGTransacciones.Row) + DBGTransacciones.RowHeight
                    .Visible = True
                    .SetFocus
                End With
             
             End If
           End If
           
           
        Set c = DBGTransacciones.Columns(ColIndex)
        With Me.TDBGridFechas
            .Left = Me.DBGTransacciones.Left + c.Left
            .top = DBGTransacciones.top + DBGTransacciones.RowTop(DBGTransacciones.Row) + DBGTransacciones.RowHeight
            .Visible = True
            .SetFocus
        End With
  
        Me.DBGTransacciones.Enabled = False

        
        End If
  
  

  
  End If




'/////Busco cambios en las claves del movimiento///////////


Select Case ColIndex
  Case 0
    '////////////Verifico la cuenta///////////////
  
'       Criterio = "CodCuentas='" & Me.DBGTransacciones.Columns(0).Text & "'"
'       Me.DtaCuentas.Recordset.Find (Criterio)
       Criterio = Me.DBGTransacciones.Columns(0).Text
       Me.AdoBuscaCuenta.RecordSource = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo From Cuentas WHERE (CodCuentas = '" & Criterio & "')"
       Me.AdoBuscaCuenta.Refresh
       


       If Not Me.AdoBuscaCuenta.Recordset.EOF Then
       
        Me.DBGTransacciones.Columns(1).Text = Me.AdoBuscaCuenta.Recordset("DescripcionCuentas")
       
 '//////////////////////////////////////////////////////////////////////////////////////////////
 '//////////SI LA CUENTA EXISTE AGREGO LOS ENCABEZADOS///////////////////////////////////////
 '/////////////////////////////////////////////////////////////////////////////////////////////
 
         FrmTransacciones.CmbMoneda.Enabled = False
         FrmTransacciones.CmdBuscarEmpleado.Enabled = False
         mes = Month(FrmTransacciones.TxtFecha.Value)
         Año = Year(FrmTransacciones.TxtFecha.Value)
         FechaIni = CDate("1/" & Month(FrmTransacciones.TxtFecha.Value) & "/" & Year(FrmTransacciones.TxtFecha.Value))
         FechaFin = DateSerial(Año, mes + 1, 1 - 1)
         NumFecha1 = FechaIni
         NumFecha2 = FechaFin
 
         FrmTransacciones.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "'))"
         FrmTransacciones.DtaConsulta.Refresh
         If Not FrmTransacciones.DtaConsulta.Recordset.EOF Then
           FrmTransacciones.TxtPeriodo.Text = FrmTransacciones.DtaConsulta.Recordset("Periodo")
            NumeroPeriodo = FrmTransacciones.DtaConsulta.Recordset("NPeriodo")
            If Val(FrmTransacciones.TxtNTransacciones.Text) = 0 Then
                NumeroTransaccion = FrmTransacciones.DtaConsulta.Recordset("NTransacciones")
            Else
                NumeroTransaccion = FrmTransacciones.TxtNTransacciones.Text
            End If
            EstadoPeriodo = FrmTransacciones.DtaConsulta.Recordset("EstadoPeriodo")
      
        '////////////Edito los datos del Periodo///////////
         If Val(FrmTransacciones.TxtNTransacciones.Text) = 0 Then
          
          
'          FrmTransacciones.'DtaConsulta.Recordset.Edit
          FrmTransacciones.DtaConsulta.Recordset("NTransacciones") = FrmTransacciones.DtaConsulta.Recordset("NTransacciones") + 1
          FrmTransacciones.DtaConsulta.Recordset.Update
          NumeroTransaccion = FrmTransacciones.DtaConsulta.Recordset("NTransacciones")
          FrmTransacciones.TxtNTransacciones.Text = NumeroTransaccion
         '////////AGREGO los Datos de los indices de Transacciones//////
         
          FrmTransacciones.DtaIndice.Recordset.AddNew
          FrmTransacciones.DtaIndice.Recordset("FechaTransaccion") = Format(FrmTransacciones.TxtFecha.Value, "dd/mm/yyyy")
          If FrmTransacciones.DBGTransacciones.Columns(3).Text <> "" Then
            FrmTransacciones.DtaIndice.Recordset("DescripcionMovimiento") = FrmTransacciones.DBGTransacciones.Columns(3).Text
          End If
          FrmTransacciones.DtaIndice.Recordset("NumeroMovimiento") = NumeroTransaccion
          FrmTransacciones.DtaIndice.Recordset("Fuente") = FrmTransacciones.TxtFuente.Text
          FrmTransacciones.DtaIndice.Recordset("NPeriodo") = NumeroPeriodo
          FrmTransacciones.DtaIndice.Recordset("TipoMoneda") = FrmTransacciones.CmbMoneda.Text
          FrmTransacciones.DtaIndice.Recordset.Update
         
        
         End If
        End If
       
       
       
       
     
       
       
       
         TipoCuenta = Me.AdoBuscaCuenta.Recordset("TipoCuenta")
         TipoMoneda = Me.AdoBuscaCuenta.Recordset("TipoMoneda")

         Select Case TipoMoneda
            Case "Córdobas"
                      Fecha = Me.TxtFecha.Value
                      Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) ='" & Format(Me.TxtFecha.Value, "yyyymmdd") & "'))"
                      Me.DtaTasas.Refresh
                If Not DtaTasas.Recordset.EOF Then
                 MontoTasa = Me.DtaTasas.Recordset("MontoCordobas")
                 If MontoTasa = 0 Then
                   MontoTasa = 1
                 End If
                 Select Case Me.CmbMoneda.Text
                  Case "Córdobas"
                    Me.DBGTransacciones.Columns(7).Text = 1
                  Case "Dólares"
                    Me.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Libras"
                    Me.DBGTransacciones.Columns(7).Text = MontoTasa
                   ' Me.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                    
                 End Select
                Else
                
                 Me.DBGTransacciones.Columns(7).Text = 1
                End If
            
            Case "Dólares"
             Fecha = Me.TxtFecha.Value
             Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) ='" & Format(Me.TxtFecha.Value, "yyyymmdd") & "'))"
             Me.DtaTasas.Refresh
             If Not DtaTasas.Recordset.EOF Then
                MontoTasa = Me.DtaTasas.Recordset("MontoCordobas")
                If MontoTasa = 0 Then
                  MontoTasa = 1
                End If
            
               Select Case Me.CmbMoneda.Text
                  Case "Córdobas"
                    Me.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                  Case "Dólares"
                    Me.DBGTransacciones.Columns(7).Text = 1
                  Case "Libras"
                    MontoTasa = Me.DtaTasas.Recordset("MontoLibras")
                    Me.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)

                    
                 End Select
                Else
                  Me.DBGTransacciones.Columns(7).Text = 1
               End If
            
            Case "Libras"
                      Fecha = Me.TxtFecha.Value
                      Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) ='" & Format(Me.TxtFecha.Value, "yyyymmdd") & "'))"
                      Me.DtaTasas.Refresh
                If Not DtaTasas.Recordset.EOF Then
                MontoTasa = Me.DtaTasas.Recordset("MontoLibras")
               Select Case Me.CmbMoneda.Text
                  Case "Córdobas"
                    Me.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Dólares"
                    Me.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Libras"
                    Me.DBGTransacciones.Columns(7).Text = 1

                    
                 End Select
                Else
                 Me.DBGTransacciones.Columns(7).Text = 1
                End If
         
         End Select
         
   TipoCuenta = Me.AdoBuscaCuenta.Recordset("TipoCuenta")
   CodigoCuenta = Me.AdoBuscaCuenta.Recordset("CodCuentas")
  If TipoCuenta = "Bancos" Or TipoCuenta = "Caja" Then

   If Primero = True Then
     Primero = False
        Me.DtaConsulta.RecordSource = "SELECT NConsecutivoVoucher.CodCuenta, NConsecutivoVoucher.ConsecutivoVoucher, NConsecutivoVoucher.NPeriodo From NConsecutivoVoucher Where (((NConsecutivoVoucher.CodCuenta) = '" & CodigoCuenta & "') And ((NConsecutivoVoucher.NPeriodo) = " & NumeroPeriodo & "))"
        Me.DtaConsulta.Refresh
        If DtaConsulta.Recordset.EOF Then
           Me.DtaConsulta.Recordset.AddNew
             Me.DtaConsulta.Recordset("CodCuenta") = CodigoCuenta
             Me.DtaConsulta.Recordset("NPeriodo") = NumeroPeriodo
             Me.DtaConsulta.Recordset("ConsecutivoVoucher") = 1
           Me.DtaConsulta.Recordset.Update
           NumeroVoucher = 1
        Else
           'Me.'DtaConsulta.Recordset.Edit
            Me.DtaConsulta.Recordset("ConsecutivoVoucher") = Me.DtaConsulta.Recordset("ConsecutivoVoucher") + 1
           Me.DtaConsulta.Recordset.Update
         NumeroVoucher = Me.DtaConsulta.Recordset("ConsecutivoVoucher")
        End If
     Else
       ' FrmCheque.DtaTransacciones.Recordset.MoveLast
     
     End If
     
     
        ConsecutivoVoucher = Month(FrmTransacciones.TxtFecha.Value)
        If TipoCuenta = "Caja" Then
              numero = "CASH " & NumeroVoucher & "/" & ConsecutivoVoucher
        End If
        Select Case TipoMoneda
           Case "Córdobas"
            If TipoCuenta = "Bancos" Then
              numero = "BC " & NumeroVoucher & "/" & ConsecutivoVoucher
            End If
           Case "Dólares"
            If TipoCuenta = "Bancos" Then
              numero = "BD " & NumeroVoucher & "/" & ConsecutivoVoucher
            End If
        
         End Select
        
     End If

   '///////////////////////////////////////////////////////////
   '//////CON ESTA CONSULTA BUSCO LA DESCRIPCION DE LA LINEA ANTERIOR//////////////////
   '/////////////////////////////////////////////////////////////////////////////////
   
            
            Sql = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta AS DescripcionCuentas, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, " & _
            "Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, " & _
            "Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, " & _
            "Transacciones.NumeroMovimiento , Periodos.Periodo " & _
            "FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo " & _
            "WHERE  (Transacciones.FechaTransaccion BETWEEN '" & Format(Me.TxtFecha.Value, "yyyymmdd") & "' And '" & Format(Me.TxtFecha.Value, "yyyymmdd") & "') AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ") " & _
            "ORDER BY Transacciones.NTransaccion"
              
            Me.DtaConsulta.RecordSource = Sql
            Me.DtaConsulta.Refresh
            If Not Me.DtaConsulta.Recordset.EOF Then
              Me.DtaConsulta.Recordset.MoveLast
              If Not IsNull(Me.DtaConsulta.Recordset("DescripcionMovimiento")) Then
                 DescripcionMovimiento = Me.DtaConsulta.Recordset("DescripcionMovimiento")
              End If
              If Not IsNull(Me.DtaConsulta.Recordset("Clave")) Then
                ClaveMovimiento = Me.DtaConsulta.Recordset("Clave")
              End If
              
              If Not IsNull(Me.DtaConsulta.Recordset("ChequeNo")) Then
                 ChequeNo = Me.DtaConsulta.Recordset("ChequeNo")
              End If
              
              
            
            End If
         
         
         Me.DBGTransacciones.Columns(2).Text = numero
         Me.DBGTransacciones.Columns(3).Text = DescripcionMovimiento
         Me.DBGTransacciones.Columns(5).Text = ChequeNo
         Me.DBGTransacciones.Columns(1).Text = Me.AdoBuscaCuenta.Recordset("DescripcionCuentas")
         Me.DBGTransacciones.Columns(10).Text = Format(Me.TxtFecha.Value, "dd/mm/yyyy")
         Me.DBGTransacciones.Columns(11).Text = NumeroPeriodo
         Me.DBGTransacciones.Columns(13).Text = Me.TxtFuente.Text
         Me.DBGTransacciones.Columns(14).Text = Format(Me.TxtFecha.Value, "dd/mm/yyyy")
         Me.DBGTransacciones.Columns(15).Text = FrmTransacciones.TxtNTransacciones.Text
         If ClaveMovimiento = "" Then
          Me.DBGTransacciones.Columns(6).Text = "Debito"
         Else
          Me.DBGTransacciones.Columns(6).Text = ClaveMovimiento
         End If
         'Me.DBGTransacciones.Columns(9).Locked = True
         'Me.DBGTransacciones.Columns(9).Locked = True
         'Me.DBGTransacciones.Columns(8).Locked = False

       Else
               
         MsgBox "La cuenta digitada no es correcta", vbCritical, "Sistema Contable"
         NumeroTransaccion = Me.TxtNTransacciones.Text
         Me.DBGTransacciones.Columns(0).Text = ""
         'Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,  Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion"
         'Me.DtaTransacciones.Refresh
            Me.DBGTransacciones.Columns(0).Button = True
              Me.DBGTransacciones.Columns(1).Locked = True
            Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
            Me.DBGTransacciones.Columns(6).Button = True
            Me.DBGTransacciones.Columns(6).Locked = True
            Me.DBGTransacciones.Columns(0).Width = 1500
            Me.DBGTransacciones.Columns(2).Width = 1100
            Me.DBGTransacciones.Columns(2).Caption = "Voucher/Dpto"
            Me.DBGTransacciones.Columns(3).Caption = "Descripcion"
            Me.DBGTransacciones.Columns(4).Width = 1000
            Me.DBGTransacciones.Columns(4).Button = True
            Me.DBGTransacciones.Columns(5).Width = 1000
            Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
            Me.DBGTransacciones.Columns(6).Width = 800
            Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
            Me.DBGTransacciones.Columns(7).Locked = True
            Me.DBGTransacciones.Columns(7).NumberFormat = "##,##0.000000"
             Me.DBGTransacciones.Columns(7).Width = 1200
            Me.DBGTransacciones.Columns(8).Width = 1200
            Me.DBGTransacciones.Columns(8).NumberFormat = "##,##0.00"
            Me.DBGTransacciones.Columns(9).Width = 1200
            Me.DBGTransacciones.Columns(9).NumberFormat = "##,##0.00"
            Me.DBGTransacciones.Columns(10).Visible = False
            Me.DBGTransacciones.Columns(11).Visible = False
            Me.DBGTransacciones.Columns(12).Visible = False
            Me.DBGTransacciones.Columns(13).Visible = False
            Me.DBGTransacciones.Columns(14).Visible = False
            Me.DBGTransacciones.Columns(15).Visible = False
              Me.DBGTransacciones.Columns(16).Visible = False
  'Me.DBGTransacciones.Enabled = False
         FrmConsulta.Show 1
         Exit Sub
       End If
     
    
 
 
       
 Case 8
   '//////////Sumo los totales del Debito///////////////
    If Me.DBGTransacciones.Columns(8).Text = "" Then
      Me.DBGTransacciones.Columns(8).Text = "0.00"
    End If
    
    Debito = Me.DBGTransacciones.Columns(8).Text
    Diferencia = Val(Debito) - Val(DebitoAnt)
    TotalDebito = TotalDebito + Diferencia
    Me.TxtDebito.Text = Format(TotalDebito, "##,##0.00")
    TotalDiferencia = TotalDebito - TotalCredito
    Me.TxtDiferencia.Text = Format(TotalDiferencia, "##,##0.00")
    
  '//////////Busco es tipo de cuenta para sumar historico///////////////////////
    CodigoCuenta = Me.DBGTransacciones.Columns(0).Text
    Criterio = "CodCuentas='" & CodigoCuenta & "'"
    Me.DtaCuentas.Recordset.Find (Criterio)
    If Not DtaCuentas.Recordset.EOF Then
     TipoCuenta = Me.DtaCuentas.Recordset("TipoCuenta")
     If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Cuentas x Pagar" Or TipoCuenta = "Cuentas de Gastos" Or TipoCuenta = "Bancos" Then
      
   '//////Busco si tiene saldo en el historial del perido actual
      Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
      Me.DtaHistorial.Refresh
       If DtaHistorial.Recordset.EOF Then
        '////Si no existe registro para este Periodo Busco los ///
        '////saldos del periodo anterior para que sean incial////
         Criterio = "NPeriodo=" & NumeroPeriodo & " "
         Me.DtaPeriodos.Recordset.Find (Criterio)
        If Not DtaPeriodos.Recordset.EOF Then
'////Busco el numero del periodo anterior para hacer la consulta
'          DtaPeriodos.Recordset.MovePrevious
'          NumeroPeriodoAnterior = DtaPeriodos.Recordset("NPeriodo")
'          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodoAnterior & "))"
'          Me.DtaHistorial.Refresh
        
          If DtaHistorial.Recordset.EOF Then
       '//////si el periodo anterior no tiene saldo////
       '/////Y el periodo actual tampoco tiene saldo/////////
           Me.DtaHistorial.Recordset.AddNew
             Me.DtaHistorial.Recordset("CodCuenta") = CodigoCuenta
             Me.DtaHistorial.Recordset("NPeriodo") = NumeroPeriodo
             Me.DtaHistorial.Recordset("SaldoInicial") = 0
             Me.DtaHistorial.Recordset("SaldoFinal") = Val(Me.DtaHistorial.Recordset("SaldoInicial")) + Diferencia
            Me.DtaHistorial.Recordset.Update
         Else
      '//////Si el periodo anterior tiene saldo////////
      '/////y el periodo actual no tiene saldo//////////
         SaldoFinal = Me.DtaHistorial.Recordset("SaldoFinal")
            Me.DtaHistorial.Recordset.AddNew
             Me.DtaHistorial.Recordset("CodCuenta") = CodigoCuenta
             Me.DtaHistorial.Recordset("NPeriodo") = NumeroPeriodo
             Me.DtaHistorial.Recordset("SaldoInicial") = SaldoFinal
             Me.DtaHistorial.Recordset("SaldoFinal") = Val(Me.DtaHistorial.Recordset("SaldoInicial")) + Diferencia
            Me.DtaHistorial.Recordset.Update
         
         End If
        End If
       Else '////////Si la cuenta tiene saldo en el periodo actual
        '////Si existe registro para este Periodo Busco los ///
        '////saldos del periodo anterior para que sean incial////
         Criterio = "NPeriodo=" & NumeroPeriodo & " "
         Me.DtaPeriodos.Recordset.Find (Criterio)
        If Not DtaPeriodos.Recordset.EOF Then
'////Busco el numero del periodo anterior para hacer la consulta
'          DtaPeriodos.Recordset.MovePrevious
'          NumeroPeriodoAnterior = DtaPeriodos.Recordset("NPeriodo")
'          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodoAnterior & "))"
'          Me.DtaHistorial.Refresh
        
          If DtaHistorial.Recordset.EOF Then
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
          Me.DtaHistorial.Refresh
          
       '//////si el periodo anterior no tiene saldo////
       '/////Y el periodo actual tiene saldo/////////
           'Me.DtaHistorial.Recordset.Edit
             Me.DtaHistorial.Recordset("SaldoInicial") = 0
             Me.DtaHistorial.Recordset("SaldoFinal") = Val(Me.DtaHistorial.Recordset("SaldoFinal")) + Diferencia
            Me.DtaHistorial.Recordset.Update
         Else
          SaldoFinal = Me.DtaHistorial.Recordset("SaldoFinal")
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
          Me.DtaHistorial.Refresh
      '//////Si el periodo anterior tiene saldo////////
      '/////y el periodo actual tiene saldo//////////
        
            'Me.DtaHistorial.Recordset.Edit
             Me.DtaHistorial.Recordset("SaldoInicial") = SaldoFinal
             Me.DtaHistorial.Recordset("SaldoFinal") = Val(Me.DtaHistorial.Recordset("SaldoFinal")) + Diferencia
            Me.DtaHistorial.Recordset.Update
         
         End If
        End If
        
        
        
       End If
     Else '///Resto el saldo actual//////////
       
       '//////Busco si tiene saldo en el historial del perido actual
      Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
      Me.DtaHistorial.Refresh
       If DtaHistorial.Recordset.EOF Then
        '////Si no existe registro para este Periodo Busco los ///
        '////saldos del periodo anterior para que sean incial////
         Criterio = "NPeriodo=" & NumeroPeriodo & " "
         Me.DtaPeriodos.Recordset.Find (Criterio)
        If Not DtaPeriodos.Recordset.EOF Then
'////Busco el numero del periodo anterior para hacer la consulta
          DtaPeriodos.Recordset.MovePrevious
          NumeroPeriodoAnterior = DtaPeriodos.Recordset("NPeriodo")
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodoAnterior & "))"
          Me.DtaHistorial.Refresh
        
          If DtaHistorial.Recordset.EOF Then
       '//////si el periodo anterior no tiene saldo////
       '/////Y el periodo actual tampoco tiene saldo/////////
           Me.DtaHistorial.Recordset.AddNew
             Me.DtaHistorial.Recordset("CodCuenta") = CodigoCuenta
             Me.DtaHistorial.Recordset("NPeriodo") = NumeroPeriodo
             Me.DtaHistorial.Recordset("SaldoInicial") = 0
             Me.DtaHistorial.Recordset("SaldoFinal") = -Diferencia
            Me.DtaHistorial.Recordset.Update
         Else
      '//////Si el periodo anterior tiene saldo////////
      '/////y el periodo actual no tiene saldo//////////
         SaldoFinal = Me.DtaHistorial.Recordset("SaldoFinal")
            Me.DtaHistorial.Recordset.AddNew
             Me.DtaHistorial.Recordset("CodCuenta") = CodigoCuenta
             Me.DtaHistorial.Recordset("NPeriodo") = NumeroPeriodo
             Me.DtaHistorial.Recordset("SaldoInicial") = SaldoFinal
             Me.DtaHistorial.Recordset("SaldoFinal") = -Diferencia
            Me.DtaHistorial.Recordset.Update
         
         End If
        End If
       Else '////////Si la cuenta tiene saldo en el periodo actual
        '////Si existe registro para este Periodo Busco los ///
        '////saldos del periodo anterior para que sean incial////
         Criterio = "NPeriodo=" & NumeroPeriodo & " "
         Me.DtaPeriodos.Recordset.Find (Criterio)
        If Not DtaPeriodos.Recordset.EOF Then
'////Busco el numero del periodo anterior para hacer la consulta
          DtaPeriodos.Recordset.MovePrevious
          NumeroPeriodoAnterior = DtaPeriodos.Recordset("NPeriodo")
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodoAnterior & "))"
          Me.DtaHistorial.Refresh
        
          If DtaHistorial.Recordset.EOF Then
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
          Me.DtaHistorial.Refresh
          
       '//////si el periodo anterior no tiene saldo////
       '/////Y el periodo actual tiene saldo/////////
           'Me.DtaHistorial.Recordset.Edit
             Me.DtaHistorial.Recordset("SaldoInicial") = 0
             Me.DtaHistorial.Recordset("SaldoFinal") = Val(Me.DtaHistorial.Recordset("SaldoFinal")) - Diferencia
            Me.DtaHistorial.Recordset.Update
         Else
          SaldoFinal = Me.DtaHistorial.Recordset("SaldoFinal")
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
          Me.DtaHistorial.Refresh
      '//////Si el periodo anterior tiene saldo////////
      '/////y el periodo actual tiene saldo//////////
        
            'Me.DtaHistorial.Recordset.Edit
             Me.DtaHistorial.Recordset("SaldoInicial") = SaldoFinal
             Me.DtaHistorial.Recordset("SaldoFinal") = Val(Me.DtaHistorial.Recordset("SaldoFinal")) - Diferencia
            Me.DtaHistorial.Recordset.Update
         
         End If
        End If
        
        
        
       End If
     
     
     
     
     
     
     End If
   End If
Case 9
    '//////////Sumo los totales del credito///////////////
    If Me.DBGTransacciones.Columns(9).Text = "" Then
      Me.DBGTransacciones.Columns(9).Text = "0.00"
    End If
    Credito = Me.DBGTransacciones.Columns(9).Text
    Diferencia = Val(Credito) - Val(CreditoAnt)
    TotalCredito = TotalCredito + Diferencia
    Me.TxtCredito.Text = Format(TotalCredito, "##,##0.00")
    TotalDiferencia = TotalDebito - TotalCredito
    Me.TxtDiferencia.Text = Format(TotalDiferencia, "##,##0.00")

  
   '//////////Busco es tipo de cuenta para sumar historico///////////////////////
    CodigoCuenta = Me.DBGTransacciones.Columns(0).Text
    Criterio = "CodCuentas='" & CodigoCuenta & "'"
    Me.DtaCuentas.Recordset.Find (Criterio)
    If Not DtaCuentas.Recordset.EOF Then
     TipoCuenta = Me.DtaCuentas.Recordset("TipoCuenta")
     If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Cuentas x Pagar" Or TipoCuenta = "Cuentas de Gastos" Or TipoCuenta = "Bancos" Then
      
   '//////Busco si tiene saldo en el historial del perido actual
      Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
      Me.DtaHistorial.Refresh
       If DtaHistorial.Recordset.EOF Then
        '////Si no existe registro para este Periodo Busco los ///
        '////saldos del periodo anterior para que sean incial////
         Criterio = "NPeriodo=" & NumeroPeriodo & " "
         Me.DtaPeriodos.Recordset.Find (Criterio)
        If Not DtaPeriodos.Recordset.EOF Then
'////Busco el numero del periodo anterior para hacer la consulta
          DtaPeriodos.Recordset.MovePrevious
          NumeroPeriodoAnterior = DtaPeriodos.Recordset("NPeriodo")
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodoAnterior & "))"
          Me.DtaHistorial.Refresh
        
          If DtaHistorial.Recordset.EOF Then
       '//////si el periodo anterior no tiene saldo////
       '/////Y el periodo actual tampoco tiene saldo/////////
           Me.DtaHistorial.Recordset.AddNew
             Me.DtaHistorial.Recordset("CodCuenta") = CodigoCuenta
             Me.DtaHistorial.Recordset("NPeriodo") = NumeroPeriodo
             Me.DtaHistorial.Recordset("SaldoInicial") = 0
             Me.DtaHistorial.Recordset("SaldoFinal") = -Diferencia
            Me.DtaHistorial.Recordset.Update
         Else
      '//////Si el periodo anterior tiene saldo////////
      '/////y el periodo actual no tiene saldo//////////
         SaldoFinal = Me.DtaHistorial.Recordset("SaldoFinal")
            Me.DtaHistorial.Recordset.AddNew
             Me.DtaHistorial.Recordset("CodCuenta") = CodigoCuenta
             Me.DtaHistorial.Recordset("NPeriodo") = NumeroPeriodo
             Me.DtaHistorial.Recordset("SaldoInicial") = SaldoFinal
             Me.DtaHistorial.Recordset("SaldoFinal") = -Diferencia
            Me.DtaHistorial.Recordset.Update
         
         End If
        End If
       Else '////////Si la cuenta tiene saldo en el periodo actual
        '////Si existe registro para este Periodo Busco los ///
        '////saldos del periodo anterior para que sean incial////
         Criterio = "NPeriodo=" & NumeroPeriodo & " "
         Me.DtaPeriodos.Recordset.Find (Criterio)
        If Not DtaPeriodos.Recordset.EOF Then
'////Busco el numero del periodo anterior para hacer la consulta
          DtaPeriodos.Recordset.MovePrevious
          NumeroPeriodoAnterior = DtaPeriodos.Recordset("NPeriodo")
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodoAnterior & "))"
          Me.DtaHistorial.Refresh
        
          If DtaHistorial.Recordset.EOF Then
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
          Me.DtaHistorial.Refresh
          
       '//////si el periodo anterior no tiene saldo////
       '/////Y el periodo actual tiene saldo/////////
           'Me.DtaHistorial.Recordset.Edit
             Me.DtaHistorial.Recordset("SaldoInicial") = 0
             Me.DtaHistorial.Recordset("SaldoFinal") = Val(Me.DtaHistorial.Recordset("SaldoFinal")) - Diferencia
            Me.DtaHistorial.Recordset.Update
         Else
          SaldoFinal = Me.DtaHistorial.Recordset("SaldoFinal")
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
          Me.DtaHistorial.Refresh
      '//////Si el periodo anterior tiene saldo////////
      '/////y el periodo actual tiene saldo//////////
        
            'Me.DtaHistorial.Recordset.Edit
             Me.DtaHistorial.Recordset("SaldoInicial") = SaldoFinal
             Me.DtaHistorial.Recordset("SaldoFinal") = Val(Me.DtaHistorial.Recordset("SaldoFinal")) - Diferencia
            Me.DtaHistorial.Recordset.Update
         
         End If
        End If
        
        
        
       End If
     Else '///Sumo el saldo//////////
       
       '//////Busco si tiene saldo en el historial del perido actual
      Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
      Me.DtaHistorial.Refresh
       If DtaHistorial.Recordset.EOF Then
        '////Si no existe registro para este Periodo Busco los ///
        '////saldos del periodo anterior para que sean incial////
         Criterio = "NPeriodo=" & NumeroPeriodo & " "
         Me.DtaPeriodos.Recordset.Find (Criterio)
        If Not DtaPeriodos.Recordset.EOF Then
'////Busco el numero del periodo anterior para hacer la consulta
'          DtaPeriodos.Recordset.MovePrevious
'          NumeroPeriodoAnterior = DtaPeriodos.Recordset("NPeriodo")
'          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodoAnterior & "))"
'          Me.DtaHistorial.Refresh
        
          If DtaHistorial.Recordset.EOF Then
       '//////si el periodo anterior no tiene saldo////
       '/////Y el periodo actual tampoco tiene saldo/////////
           Me.DtaHistorial.Recordset.AddNew
             Me.DtaHistorial.Recordset("CodCuenta") = CodigoCuenta
             Me.DtaHistorial.Recordset("NPeriodo") = NumeroPeriodo
             Me.DtaHistorial.Recordset("SaldoInicial") = 0
             Me.DtaHistorial.Recordset("SaldoFinal") = Diferencia
            Me.DtaHistorial.Recordset.Update
         Else
      '//////Si el periodo anterior tiene saldo////////
      '/////y el periodo actual no tiene saldo//////////
         SaldoFinal = Me.DtaHistorial.Recordset("SaldoFinal")
            Me.DtaHistorial.Recordset.AddNew
             Me.DtaHistorial.Recordset("CodCuenta") = CodigoCuenta
             Me.DtaHistorial.Recordset("NPeriodo") = NumeroPeriodo
             Me.DtaHistorial.Recordset("SaldoInicial") = SaldoFinal
             Me.DtaHistorial.Recordset("SaldoFinal") = Diferencia
            Me.DtaHistorial.Recordset.Update
         
         End If
        End If
       Else '////////Si la cuenta tiene saldo en el periodo actual
        '////Si existe registro para este Periodo Busco los ///
        '////saldos del periodo anterior para que sean incial////
         Criterio = "NPeriodo=" & NumeroPeriodo & " "
         Me.DtaPeriodos.Recordset.Find (Criterio)
        If Not DtaPeriodos.Recordset.EOF Then
'////Busco el numero del periodo anterior para hacer la consulta
'          DtaPeriodos.Recordset.MovePrevious
'          NumeroPeriodoAnterior = DtaPeriodos.Recordset("NPeriodo")
'          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodoAnterior & "))"
'          Me.DtaHistorial.Refresh
        
          If DtaHistorial.Recordset.EOF Then
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
          Me.DtaHistorial.Refresh
          
       '//////si el periodo anterior no tiene saldo////
       '/////Y el periodo actual tiene saldo/////////
           'Me.DtaHistorial.Recordset.Edit
             Me.DtaHistorial.Recordset("SaldoInicial") = 0
             Me.DtaHistorial.Recordset("SaldoFinal") = Val(Me.DtaHistorial.Recordset("SaldoFinal")) + Diferencia
            Me.DtaHistorial.Recordset.Update
         Else
          SaldoFinal = Me.DtaHistorial.Recordset("SaldoFinal")
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
          Me.DtaHistorial.Refresh
      '//////Si el periodo anterior tiene saldo////////
      '/////y el periodo actual tiene saldo//////////
        
            'Me.DtaHistorial.Recordset.Edit
             Me.DtaHistorial.Recordset("SaldoInicial") = SaldoFinal
             Me.DtaHistorial.Recordset("SaldoFinal") = Val(Me.DtaHistorial.Recordset("SaldoFinal")) + Diferencia
            Me.DtaHistorial.Recordset.Update
         
         End If
        End If
        
        
        
       End If
     
     
     
     
     
     
     End If
   End If
End Select

Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub DBGTransacciones_AfterColUpdate(ByVal ColIndex As Integer)
On Error GoTo TipoErrs
Dim Fechas1 As String, Fechas2 As String
   Select Case ColIndex
    
    Case 0
    
    
      mes = Month(Me.TxtFecha.Value)
      Año = Year(Me.TxtFecha.Value)
      FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
      FechaFin = DateSerial(Año, mes + 1, 1 - 1)
      NumFecha1 = FechaIni
      NumFecha2 = FechaFin
      Fechas1 = FechaIni
      Fechas2 = FechaFin
      
      Me.DtaConsulta.RecordSource = "SELECT NPeriodo, NumeroTabla, FechaPeriodo, EstadoPeriodo, NTransacciones, Periodo From Periodos WHERE     (FechaPeriodo BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas1 & "', 102))"
      '      Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
      Me.DtaConsulta.Refresh
      If Not DtaConsulta.Recordset.EOF Then
        Me.TxtPeriodo.Text = DtaConsulta.Recordset("Periodo")
        NumeroPeriodo = DtaConsulta.Recordset("NPeriodo")
        If Val(Me.TxtNTransacciones.Text) = 0 Then
         NumeroTransaccion = DtaConsulta.Recordset("NTransacciones")
        End If
        EstadoPeriodo = DtaConsulta.Recordset("EstadoPeriodo")
      
      '////////////Edito los datos del Periodo///////////
     If Val(Me.TxtNTransacciones.Text) = 0 Then
     
     
     
     
      'Me.'DtaConsulta.Recordset.Edit
        DtaConsulta.Recordset("NTransacciones") = DtaConsulta.Recordset("NTransacciones") + 1
      Me.DtaConsulta.Recordset.Update
      NumeroTransaccion = DtaConsulta.Recordset("NTransacciones")
      FrmTransacciones.TxtNTransacciones.Text = NumeroTransaccion
      '////////Edito los Datos de los indices de Transacciones//////
         
          Me.DtaIndice.Recordset.AddNew
          Me.DtaIndice.Recordset("FechaTransaccion") = Format(Me.TxtFecha.Value, "dd/mm/yyyy")
          Me.DtaIndice.Recordset("NumeroMovimiento") = NumeroTransaccion
          Me.DtaIndice.Recordset("DescripcionMovimiento") = Me.DBGTransacciones.Columns(1).Text
          Me.DtaIndice.Recordset("Fuente") = Me.TxtFuente.Text
          Me.DtaIndice.Recordset("NPeriodo") = NumeroPeriodo
          Me.DtaIndice.Recordset("TipoMoneda") = Me.CmbMoneda.Text
          Me.DtaIndice.Recordset.Update
      
      
      
     
     
     
     End If
   End If

     
   Case 3
      mes = Month(Me.TxtFecha.Value)
      Año = Year(Me.TxtFecha.Value)
      FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
      FechaFin = DateSerial(Año, mes + 1, 1 - 1)
      NumFecha1 = FechaIni
      NumFecha2 = FechaFin
      Fechas1 = FechaIni
      Fechas2 = FechaFin
      
      Me.DtaConsulta.RecordSource = "SELECT NPeriodo, NumeroTabla, FechaPeriodo, EstadoPeriodo, NTransacciones, Periodo From Periodos WHERE     (FechaPeriodo BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas1 & "', 102))"
      Me.DtaConsulta.Refresh
      If Not DtaConsulta.Recordset.EOF Then
        NumeroPeriodo = DtaConsulta.Recordset("NPeriodo")
        If Val(Me.TxtNTransacciones.Text) = 0 Then
         NumeroTransaccion = DtaConsulta.Recordset("NTransacciones")
        Else
         NumeroTransaccion = Me.TxtNTransacciones.Text
        End If
        EstadoPeriodo = DtaConsulta.Recordset("EstadoPeriodo")
      End If
  Me.AdoBuscar.RecordSource = "SELECT FechaTransaccion, NumeroMovimiento, Nperiodo, DescripcionMovimiento, Fuente, TipoMoneda From IndiceTransaccion Where (NPeriodo = " & NumeroPeriodo & ") And (NumeroMovimiento = " & NumeroTransaccion & ")"
  Me.AdoBuscar.Refresh
  
   If Not Me.AdoBuscar.Recordset.EOF Then
   Me.AdoBuscar.Recordset("DescripcionMovimiento") = Me.DBGTransacciones.Columns(3).Text
   Me.AdoBuscar.Recordset.Update
   End If
     
   Case 4
   
  
      If Me.ChkVentana.Value = 1 Then
  
              '//////////////////////////////////////////////////////////
              '////BUSCO EL TIPO DE CUENTA PARA MOSTRAR FECHAS///////////
              '////////////////////////////////////////////////////////
              
               CodigoCuenta = Me.DBGTransacciones.Columns(0).Text
               Me.DtaConsulta.RecordSource = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo From Cuentas WHERE (CodCuentas = '" & CodigoCuenta & "')"
               Me.DtaConsulta.Refresh
               If Not Me.DtaConsulta.Recordset.EOF Then
                 TipoCuenta = Me.DtaConsulta.Recordset("TipoCuenta")
                 If TipoCuenta = "Cuentas x Cobrar" Then
                    Set c = DBGTransacciones.Columns(ColIndex)
                    With Me.TDBGridFechas
                        .Left = Me.DBGTransacciones.Left + c.Left
                        .top = DBGTransacciones.top + DBGTransacciones.RowTop(DBGTransacciones.Row) + DBGTransacciones.RowHeight
                        .Visible = True
                        .SetFocus
            
                    End With
                 ElseIf TipoCuenta = "Cuentas x Pagar" Then
                    Set c = DBGTransacciones.Columns(ColIndex)
                    With Me.TDBGridFechas
                        .Left = Me.DBGTransacciones.Left + c.Left
                        .top = DBGTransacciones.top + DBGTransacciones.RowTop(DBGTransacciones.Row) + DBGTransacciones.RowHeight
                        .Visible = True
                        .SetFocus
                    End With
                 
                 End If
               End If
            
                    Set c = DBGTransacciones.Columns(ColIndex)
                    With Me.TDBGridFechas
                        .Left = Me.DBGTransacciones.Left + c.Left
                        .top = DBGTransacciones.top + DBGTransacciones.RowTop(DBGTransacciones.Row) + DBGTransacciones.RowHeight
                        .Visible = True
                        .SetFocus
                    End With
  
       End If
  
  End Select
  
  Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub DBGTransacciones_AfterUpdate()
Dim Fechas As String




Salir = False
 Debito = 0
 Credito = 0
 TotalDebito = 0
 TotalCredito = 0
      NumFecha1 = Me.TxtFecha.Value
      Fechas = Format(Me.TxtFecha.Value, "yyyy/mm/dd")
      NMovimiento = Val(Me.TxtNTransacciones)
      Me.DtaConsulta.RecordSource = "SELECT FechaTransaccion, CodCuentas, NTransaccion, NumeroMovimiento, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, TCambio, Debito, Credito From Transacciones WHERE     (FechaTransaccion = CONVERT(DATETIME, '" & Fechas & "', 102)) AND (NumeroMovimiento = " & NMovimiento & ")"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
       Debito = Me.DtaConsulta.Recordset("Debito")
       Credito = Me.DtaConsulta.Recordset("Credito")
       TotalDebito = TotalDebito + Debito
       TotalCredito = TotalCredito + Credito
       DtaConsulta.Recordset.MoveNext
      Loop
Me.DBGTransacciones.Columns(7).NumberFormat = "##,##0.000000"
Me.TxtCredito.Text = Format(TotalCredito, "##,##0.00")
Me.TxtDebito.Text = Format(TotalDebito, "##,##0.00")
Me.TxtDiferencia.Text = Format(TotalDebito - TotalCredito, "##,##0.00")

Me.DBGTransacciones.PostMsg (2)
        
Me.CmdNuevo.Enabled = False
End Sub

Private Sub DBGTransacciones_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
On Error GoTo TipoErrs
If ColIndex = 8 Or ColIndex = 9 Then
 If Me.DBGTransacciones.Columns(6).Text = "Debito" Then
    If Not ColIndex = 8 Then
      Me.DBGTransacciones.Columns(ColIndex) = "0.00"
    End If
      'Me.DBGTransacciones.Columns(9).Locked = True
      'Me.DBGTransacciones.Columns(8).Locked = False
  ElseIf Me.DBGTransacciones.Columns(6).Text = "Credito" Then
      If Not ColIndex = 9 Then
        Me.DBGTransacciones.Columns(ColIndex) = "0.00"
      End If
     'Me.DBGTransacciones.Columns(9).Locked = False
     'Me.DBGTransacciones.Columns(8).Locked = True
 
 End If
 '///////Guardo la Clave Anterior//////////
 If Not Me.DBGTransacciones.Columns(6).Text = "" Then
  ClaveAnt = Me.DBGTransacciones.Columns(6).Text
 Else
 ClaveAnt = 0
 Me.DBGTransacciones.Columns(6).Text = "Debito"
 End If
 
 
 
 '///////Guardo el Debito anterior//////
 If Not Me.DBGTransacciones.Columns(8).Text = "" Then
  DebitoAnt = Me.DBGTransacciones.Columns(8).Text
 Else
 DebitoAnt = 0
 Me.DBGTransacciones.Columns(8).Text = "0.00"
 End If
 '////////////Guardo el credito anterior////////
 If Not Me.DBGTransacciones.Columns(9).Text = "" Then
  CreditoAnt = Me.DBGTransacciones.Columns(9).Text
 Else
 DebitoAnt = 0
 Me.DBGTransacciones.Columns(9).Text = "0.00"
 End If
End If




Exit Sub
TipoErrs:
 ControlErrores
End Sub



Private Sub DBGTransacciones_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
 Dim Criterio As String
 '/////////////////REVALIDO SI LA CUENTA EXISTE /////////////////////////////////
       Criterio = Me.DBGTransacciones.Columns(0).Text
       Me.AdoBuscaCuenta.RecordSource = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo From Cuentas WHERE (CodCuentas = '" & Criterio & "')"
       Me.AdoBuscaCuenta.Refresh
       If Not Me.AdoBuscaCuenta.Recordset.EOF Then
         Me.DBGTransacciones.Columns(1).Text = Me.AdoBuscaCuenta.Recordset("DescripcionCuentas")
         Select Case ColIndex
            Case 2
                If Me.DBGTransacciones.Columns(2).Text <> "" And Me.DBGTransacciones.Columns(2).Text <> "-" Then
                 If ExisteDpto(Me.DBGTransacciones.Columns(2).Text) = False Then
                   MsgBox "No Existe este Departamento", vbCritical, "Zeus contable"
                   Me.DBGTransacciones.Columns(2).Text = ""
                 End If
                End If
            End Select
       Else
'         MsgBox "No Existe la Cuenta", vbCritical, "Zeus Facturacion"
         Me.DBGTransacciones.Columns(0).Text = ""
         Me.DBGTransacciones.Columns(1).Text = ""
       End If
       
       
       
End Sub

Private Sub DBGTransacciones_BeforeUpdate(Cancel As Integer)
On Error GoTo TipoErrs

If Me.DBGTransacciones.Columns(6).Text = "" Then
   Me.DBGTransacciones.Columns(6).Text = "Debito"
 End If
 
 If Me.DBGTransacciones.Columns(8).Text = "" Then
   Me.DBGTransacciones.Columns(8).Text = 0
 End If
 If Me.DBGTransacciones.Columns(9).Text = "" Then
   Me.DBGTransacciones.Columns(9).Text = 0
 End If
 
 If Me.DBGTransacciones.Columns(6).Text = "Debito" Then
  Me.DBGTransacciones.Columns(9).Text = 0#
 End If
 
 If Me.DBGTransacciones.Columns(6).Text = "Credito" Then
  Me.DBGTransacciones.Columns(8).Text = 0#
 End If
  For i = 2 To 5
            If Me.DBGTransacciones.Columns(i).Text = "" Then DBGTransacciones.Columns(i).Text = "-"
        Next i
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub DBGTransacciones_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo TipoErrs
QUIEN = "No"
Select Case ColIndex
  Case 0
  
  QueProducto = "Transacciones"
  FrmConsulta.Show 1
  Case 2
   QueProducto = "Departamento"
   FrmConsulta.Show 1
   Me.DBGTransacciones.Columns(2).Text = FrmConsulta.Codigo
  
  Case 6
    Set c = DBGTransacciones.Columns(ColIndex)
      With List1
      .Left = Me.DBGTransacciones.Left + c.Left
      .top = DBGTransacciones.top + DBGTransacciones.RowTop(DBGTransacciones.Row) + DBGTransacciones.RowHeight
      .Width = c.Width + 15
      .Visible = True
      .SetFocus
      '.BoundText = Descripcion
      End With
  Case 4
  
  
  
     FrmTransaccionCob.Show

End Select

Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub DBGTransacciones_FilterChange()
On Error GoTo TipoErrs:
Dim Filtro As String
Set cols = DBGTransacciones.Columns
Dim c As Integer
c = DBGTransacciones.col
DBGTransacciones.HoldFields
Filtro = getFilter()
DtaTransacciones.Recordset.Filter = Filtro
DBGTransacciones.col = c
DBGTransacciones.EditActive = True

Exit Sub
TipoErrs:
 MsgBox err.Description
End Sub

Function getFilter() As String

Dim tmp As String
Dim n As Integer
For Each col In cols
If Trim(col.FilterText) <> "" Then
n = n + 1
If n > 1 Then
tmp = tmp & " AND "
End If
tmp = tmp & col.DataField & " LIKE '" & col.FilterText & "*'"
End If
Next col

getFilter = tmp
End Function

Private Sub DBGTransacciones_GotFocus()
On Error GoTo TipoErrs
Dim Fechas1 As String, Fechas2 As String
 mes = Month(Me.TxtFecha.Value)
      Año = Year(Me.TxtFecha.Value)
      FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
      FechaFin = DateSerial(Año, mes + 1, 1 - 1)
      NumFecha1 = FechaIni
      NumFecha2 = FechaFin
      Fechas1 = Format(FechaIni, "yyyy/mm/dd")
      Fechas2 = Format(FechaFin, "yyyy/mm/dd")
      
      Me.DtaConsulta.RecordSource = "SELECT NPeriodo, NumeroTabla, FechaPeriodo, EstadoPeriodo, NTransacciones, Periodo From Periodos WHERE     (FechaPeriodo BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas2 & "', 102))"
  
 
'      Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
      Me.DtaConsulta.Refresh
      If Not DtaConsulta.Recordset.EOF Then
        Me.TxtPeriodo.Text = DtaConsulta.Recordset("Periodo")
        NumeroPeriodo = DtaConsulta.Recordset("NPeriodo")
        If Val(Me.TxtNTransacciones.Text) = 0 Then
        NumeroTransaccion = DtaConsulta.Recordset("NTransacciones")
        End If
        EstadoPeriodo = DtaConsulta.Recordset("EstadoPeriodo")
        If EstadoPeriodo = "B" Then
           MsgBox "El Periodo Esta Bloqueado", vbCritical, "Sistema Contable"
           'Me.TxtFecha.SetFocus
           TxtFecha.Enabled = True
           Me.TxtPeriodo.Enabled = True
           Me.TxtFuente.Enabled = True
           Me.TxtNTransacciones.Enabled = True
           Me.DBGTransacciones.Enabled = False
           Exit Sub
        ElseIf EstadoPeriodo = "C" Then
           MsgBox "El Periodo esta Cerrado", vbCritical, "Sistema Contable"
           Me.TxtFecha.SetFocus
           TxtFecha.Enabled = True
           Me.TxtPeriodo.Enabled = True
           Me.TxtFuente.Enabled = True
           Me.TxtNTransacciones.Enabled = True
           Me.DBGTransacciones.Enabled = False
           Exit Sub
        Else
           Me.DBGTransacciones.Enabled = True
        End If
            
      Else
        MsgBox "La Fecha esta fuera del Rango de Periodos", vbCritical, "Sistema Contable"
        Me.DBGTransacciones.Enabled = False
        TxtFecha.Enabled = True
        Me.TxtPeriodo.Enabled = True
        Me.TxtFuente.Enabled = True
        Me.TxtNTransacciones.Enabled = True
        Exit Sub
      End If



 TxtFecha.Enabled = False
 Me.TxtPeriodo.Enabled = False
 Me.TxtFuente.Enabled = False
 Me.TxtNTransacciones.Enabled = False
 
 
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub DBGTransacciones_KeyDown(KeyCode As Integer, Shift As Integer)
Dim NumeroMovimiento As Double, Fecha As String, NTransaccion As Double, CodCuenta As String, Indice As Double
On Error GoTo TipoErrs
QUIEN = ""
Indice = 3
Select Case KeyCode
Case 37
       If Me.DBGTransacciones.col = 7 Then
            Set c = DBGTransacciones.Columns(6)
              Select Case DBGTransacciones.Columns(6).Text
                 Case "Debito"
                    Indice = 0
                 Case "Credito"
                    Indice = 1
                 Case Else
                    Indice = 3
              End Select
              
              QUIEN = "Grid"
              With List1
              .Left = Me.DBGTransacciones.Left + c.Left
              .top = DBGTransacciones.top + DBGTransacciones.RowTop(DBGTransacciones.Row) + DBGTransacciones.RowHeight
              .Width = c.Width + 15
              .Visible = True
              .SetFocus
              .Selected(Indice) = True
              End With
              QUIEN = ""
       End If
Case 39
       If Me.DBGTransacciones.col = 5 Then
            Set c = DBGTransacciones.Columns(6)
              Select Case DBGTransacciones.Columns(6).Text
                 Case "Debito"
                    Indice = 0
                 Case "Credito"
                    Indice = 1
                 Case Else
                    Indice = 3
              End Select
            
              QUIEN = "Grid"
              With List1
              .Left = Me.DBGTransacciones.Left + c.Left
              .top = DBGTransacciones.top + DBGTransacciones.RowTop(DBGTransacciones.Row) + DBGTransacciones.RowHeight
              .Width = c.Width + 15
              .Visible = True
              .SetFocus
              .Selected(Indice) = True
              End With
              QUIEN = ""
       End If
Case 13
       If Me.DBGTransacciones.col = 9 Then
         Me.DBGTransacciones.PostMsg (2)
       End If
 Case 113
        QueProducto = "Transacciones"
        FrmConsulta.Show 1

 
  Case 114
        Indice = 1
           
        Criterio = "CodCuentas='" & Me.DBGTransacciones.Columns(0).Text & "'"
        Me.DtaCuentas.Recordset.Find (Criterio)
        If Not DtaCuentas.Recordset.EOF Then
           TipoMoneda = DtaCuentas.Recordset("TipoMoneda")
        End If
         FrmConvertir.LblNombre.Caption = "Monto " & TipoMoneda
         FrmConvertir.TxtTasa.Text = Me.DBGTransacciones.Columns(7).Text
         
         FrmConvertir.Show 1
  
  Case 123
    If Not Me.DBGTransacciones.Columns(4).Text = "" Then
     If Not Me.DBGTransacciones.Columns(4).Text = "-" Then
       
       NumeroMovimiento = Me.TxtNTransacciones.Text
       Fecha = Format(Me.TxtFecha.Value, "yyyy-mm-dd")
       NTransaccion = Me.DBGTransacciones.Columns(22).Text
       Me.AdoBuscar.RecordSource = "SELECT * From Transacciones WHERE (FechaTransaccion = CONVERT(DATETIME, '" & Fecha & "', 102)) AND (NumeroMovimiento = " & NumeroMovimiento & ") AND (NTransaccion = " & NTransaccion & ")"
       Me.AdoBuscar.Refresh
       If Not Me.AdoBuscar.Recordset.EOF Then
         Me.DTPFechaCredito.Value = Me.AdoBuscar.Recordset("FechaDescuento")
         Me.DTPFechaVence.Value = Me.AdoBuscar.Recordset("FechaVence")
         Me.TxtMonto.Text = Me.AdoBuscar.Recordset("DescuentoDisponible")
         If Not IsNull(Me.AdoBuscar.Recordset("CodCuentaProveedor")) Then
          CodCuentas = Me.AdoBuscar.Recordset("CodCuentaProveedor")
         End If
         
         If Me.AdoBuscar.Recordset("TipoFactura") = "FacturaCompra" Then
           Me.OptFacturaCompra.Value = True
         ElseIf Me.AdoBuscar.Recordset("TipoFactura") = "FacturaVenta" Then
           Me.OptFacturaVenta.Value = True
         End If
         
       End If
       
       
       Me.AdoBuscaCuenta.RecordSource = "SELECT * From Cuentas WHERE (CodCuentas = '" & CodCuentas & "')"
       Me.AdoBuscaCuenta.Refresh
       If Not Me.AdoBuscaCuenta.Recordset.EOF Then
        Me.LblNombres.Caption = Me.AdoBuscaCuenta.Recordset("DescripcionCuentas")
        
        Me.TDBProveedor.Text = CodCuentas
       
         If Me.AdoBuscaCuenta.Recordset("CausaIva") = True Then
          Me.OptIva.Value = True
         ElseIf Me.AdoBuscaCuenta.Recordset("CausaRetencion") = True Then
          Me.OptRetencion.Value = True
         End If
       
       End If
       
       Me.DBGTransacciones.Enabled = False
       
       Set c = DBGTransacciones.Columns(4)
        With Me.TDBGridFechas
            .Left = Me.DBGTransacciones.Left + c.Left
            .top = DBGTransacciones.top + DBGTransacciones.RowTop(DBGTransacciones.Row) + DBGTransacciones.RowHeight
            .Visible = True
            .SetFocus
        End With
       
       
       
     End If
    End If

 End Select
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 

End If
End Sub


Private Sub DBGTransacciones_PostEvent(ByVal MsgId As Integer)

   Select Case MsgId
       Case 1
             DBGTransacciones.Refresh
       Case 2
'            DBGTransacciones.SetFocus
            'Set focus to split zero and column 0
            DBGTransacciones.Split = 0
            DBGTransacciones.col = 0
       Case 3
            DBGTransacciones.SetFocus
            'Set focus to split zero and column 0
            DBGTransacciones.Split = 0
            DBGTransacciones.col = 8
      Case 4
            DBGTransacciones.SetFocus
            'Set focus to split zero and column 0
            DBGTransacciones.Split = 0
            DBGTransacciones.col = 9
   
   End Select

    
End Sub

Private Sub DBGTransacciones_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

Select Case Me.DBGTransacciones.col
  Case 1:
   Me.DBGTransacciones.Columns(1).Width = 3000
   Me.DBGTransacciones.Columns(2).Button = False
   Me.DBGTransacciones.Columns(3).Caption = "Descripcion"
   Me.DBGTransacciones.Columns(3).Width = 2000
  Case 2:
   If MostrarBotonDpto(Me.DBGTransacciones.Columns(0).Text) = True Then
     Me.DBGTransacciones.Columns(2).Button = True
    If Me.DBGTransacciones.Columns(2).Text <> "" And Me.DBGTransacciones.Columns(2).Text <> "-" Then
     If ExisteDpto(Me.DBGTransacciones.Columns(2).Text) = False Then
       MsgBox "No Existe este Departamento", vbCritical, "Zeus contable"
       Me.DBGTransacciones.Columns(2).Text = ""
     End If
    End If
   Else
     Me.DBGTransacciones.Columns(2).Button = False
   End If
   Me.DBGTransacciones.Columns(3).Caption = "Descripcion"
   Me.DBGTransacciones.Columns(3).Width = 2000
   Me.DBGTransacciones.Columns(1).Width = 1500
    
  Case 3:
    Me.DBGTransacciones.Columns(3).Caption = "Descripcion"
    Me.DBGTransacciones.Columns(3).Width = 5000
    Me.DBGTransacciones.Columns(1).Width = 1500
    Me.DBGTransacciones.Columns(2).Button = False

  Case Else
   Me.DBGTransacciones.Columns(2).Button = False
   Me.DBGTransacciones.Columns(3).Caption = "Descripcion"
   Me.DBGTransacciones.Columns(3).Width = 2000
   Me.DBGTransacciones.Columns(1).Width = 1500
End Select

End Sub

Private Sub DTPFechaCredito_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  TxtMonto.SetFocus

End If
End Sub

Private Sub DTPFechaVence_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Fecha As Date, Fechas As String
 If KeyCode = 13 Then
  TDBGridFechas.Visible = False


FechaFactura = Format(Me.DTPFechaCredito.Value, "dd/mm/yyyy")
FechaVence = Format(Me.DTPFechaVence.Value, "dd/mm/yyyy")
Monto = Val(Me.TxtMonto.Text)


Me.DBGTransacciones.Columns(17).Text = FechaFactura
Me.DBGTransacciones.Columns(18).Text = Monto
Me.DBGTransacciones.Columns(19).Text = FechaVence
 
 End If


End Sub

Private Sub Form_Activate()
On Error GoTo TipoErrs
 

 
If Not CodigoUsuario = 0 Then
Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Transacciones'))"
Me.DtaNacceso.Refresh
If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdGrabar.Enabled = False
   Me.TxtFecha.Enabled = False
   Me.Frame1.Enabled = False
   Me.DBGTransacciones.Enabled = False
   Me.CmdBuscarEmpleado.Enabled = False
Else
  Me.DBGTransacciones.Enabled = True
End If
Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Transacciones'))"
Me.DtaNacceso.Refresh
If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdBorrar.Enabled = False
   Me.SmartButton1.Enabled = False
End If
End If
Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub Form_Load()
On Error GoTo TipoErrs
Dim Sql As String
Dim Fechas1 As String, Fechas2 As String

Me.TxtFecha.Value = Format(FechaSistema, "dd/mm/yyyy")

MDIPrimero.Skin1.ApplySkin hWnd
Salir = True
TotalDebito = 0
TotalCredito = 0
Debito = 0
Credito = 0

Me.DTPFechaCredito.Value = Format(FechaSistema, "dd/mm/yyyy")
Me.DTPFechaVence.Value = Format(FechaSistema, "dd/mm/yyyy")

 Me.DBGTransacciones.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DBGTransacciones.OddRowStyle.BackColor = &H80000005
 Me.DBGTransacciones.AlternatingRowStyle = True
 Me.TDBGridFechas.BackColor = RGB(216, 228, 248)
 Me.Label1.BackColor = RGB(216, 228, 248)
 Me.Label2.BackColor = RGB(216, 228, 248)
 Me.Label3.BackColor = RGB(216, 228, 248)
 Me.Label4.BackColor = RGB(216, 228, 248)
 Me.LblProveedor.BackColor = RGB(216, 228, 248)
 Me.LblNombres.BackColor = RGB(216, 228, 248)
 Me.GroupBox1.BackColor = RGB(216, 228, 248)
 Me.GroupBox2.BackColor = RGB(216, 228, 248)

Primero = True
With Me.DtaNacceso
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Accesos"
   .Refresh
End With

With Me.AdoProveedores
   .ConnectionString = Conexion
End With

With Me.AdoFuente
   .ConnectionString = Conexion
End With

With Me.AdoBuscar
   .ConnectionString = Conexion
End With

With Me.AdoFechasVence
   .ConnectionString = Conexion
End With

With Me.DtaPeriodos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Periodos"
   .Refresh
End With

With Me.DtaHistorial
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaIndice
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "IndiceTransaccion"
   .Refresh
End With

With Me.DtaTransaccionesNuevas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaCuentas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Cuentas"
   .Refresh
End With


With Me.DtaTasas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaConsulta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaTransacciones
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.AdoBuscaCuenta
   .ConnectionString = Conexion
End With


'Set RstFechas = New ADODB.Recordset
'With RstFechas
'     .Fields.Append "FechaDescuento", adDate
'     .Fields.Append "MontoDescuento", adDouble, 20
'     .Fields.Append "FechaVence", adDate
'     .Open
'End With

'Set Me.TDBGridFechas.DataSource = RstFechas
'    Me.TDBGridFechas.ReBind


Me.AdoFuente.RecordSource = "SELECT DISTINCT Fuente From IndiceTransaccion WHERE (Fuente <> N'  ')"
Me.AdoFuente.Refresh

' Me.TxtFuente.Clear
'Do While Not Me.AdoFuente.Recordset.EOF
' Me.TxtFuente.AddItem Me.AdoFuente.Recordset("Fuente")
'
' Me.AdoFuente.Recordset.MoveNext
'Loop

'/////////////////////////////////////////////////////////////////////
'/////////////BUSCO EL PERIODO PARA LA FECHA////////////////
'//////////////////////////////////////////////////////////////

 Me.DBGTransacciones.Enabled = True
 mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 Fechas1 = Format(FechaIni, "yyyy/mm/dd")
 Fechas2 = Format(FechaFin, "yyyy/mm/dd")
 
 Me.DtaConsulta.RecordSource = "SELECT NPeriodo, NumeroTabla, FechaPeriodo, EstadoPeriodo, NTransacciones, Periodo From Periodos WHERE     (FechaPeriodo BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas2 & "', 102))"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  Me.TxtPeriodo.Text = DtaConsulta.Recordset("Periodo")
  NumeroPeriodo = DtaConsulta.Recordset("NPeriodo")
  NumeroTransaccion = DtaConsulta.Recordset("NTransacciones")
  EstadoPeriodo = DtaConsulta.Recordset("EstadoPeriodo")
  If EstadoPeriodo = "B" Then
   MsgBox "El Periodo Esta Bloqueado", vbCritical, "Sistema Contable"
'   Me.TxtFecha.SetFocus
   Me.DBGTransacciones.Enabled = False
   TxtFecha.Enabled = True
   Me.TxtPeriodo.Enabled = True
   Me.TxtFuente.Enabled = True
   Me.TxtNTransacciones.Enabled = True
'   Exit Sub
  ElseIf EstadoPeriodo = "C" Then
  MsgBox "El Periodo esta Cerrado", vbCritical, "Sistema Contable"
'  Me.TxtFecha.SetFocus
  TxtFecha.Enabled = True
  Me.TxtPeriodo.Enabled = True
  Me.TxtFuente.Enabled = True
  Me.TxtNTransacciones.Enabled = True
  Me.DBGTransacciones.Enabled = False
'  Exit Sub
  Else
   Me.DBGTransacciones.Enabled = True
  If Not CodigoUsuario = 0 Then
   Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Transacciones'))"
   Me.DtaNacceso.Refresh
   If Me.DtaNacceso.Recordset.EOF Then
    Me.DBGTransacciones.Enabled = False
   Else
     Me.DBGTransacciones.Enabled = True
   End If
      
  End If
  End If
 Else
   MsgBox "La Fecha esta fuera del Rango de Periodos", vbCritical, "Sistema Contable"
   Me.DBGTransacciones.Enabled = False
   TxtFecha.Enabled = True
   Me.TxtPeriodo.Enabled = True
   Me.TxtFuente.Enabled = True
   Me.TxtNTransacciones.Enabled = True
'   Exit Sub
 End If
 
 '///////Verifico si esta registrada la fecha de la tasa//////

NumFecha = Me.TxtFecha.Value
Fechas = Format(Me.TxtFecha.Value, "yyyy/mm/dd")
Me.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas & "', 102))ORDER BY FechaTasas"
'DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) = " & NumFecha & "))ORDER BY Tasas.FechaTasas"
DtaTasas.Refresh

If Not DtaTasas.Recordset.EOF Then
Fecha = Format(DtaTasas.Recordset("FechaTasas"), "dd/mm/yyyy")
   
    Encontrado = True
    Cambio = DtaTasas.Recordset("MontoCordobas")
    MDIPrimero.StatusBar2.Panels(2) = "Tasa Dolar: " & Format(Cambio, "##,##0.00") & "    " & "Tasa Libras: " & Format(DtaTasas.Recordset("MontoLibras"), "##,##0.00")
End If

If Not Encontrado Then
  MsgBox "La tasa de esta Fecha no ha sido Grabada"
  Cancel = 100
  Tasa = False
  frmTasa2.Show 1
End If


Sql = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, " & _
       "Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, " & _
       "Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, " & _
       "Transacciones.NumeroMovimiento, Periodos.Periodo, Transacciones.FechaDescuento, Transacciones.DescuentoDisponible, " & _
       "Transacciones.FechaVence,Transacciones.CodCuentaProveedor,Transacciones.TipoFactura " & _
       "FROM         Periodos INNER JOIN " & _
       "Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo " & _
       "Where (Transacciones.NumeroMovimiento = -1) " & _
       "ORDER BY Transacciones.NTransaccion "
       
Me.DtaTransacciones.RecordSource = Sql
Me.DtaTransacciones.Refresh
Me.CmbMoneda.Text = "Córdobas"
  
  Me.DBGTransacciones.Columns("CodCuentas").Width = 1500  '0
  Me.DBGTransacciones.Columns("CodCuentas").Button = True  '0
  Me.DBGTransacciones.Columns("NombreCuenta").Locked = True '1
  Me.DBGTransacciones.Columns("NombreCuenta").Locked = True '1
  Me.DBGTransacciones.Columns("VoucherNo").Width = 1000 '2
  Me.DBGTransacciones.Columns("VoucherNo").Caption = "Voucher/Dpto" '2
  Me.DBGTransacciones.Columns("VoucherNo").Width = 1100 '2
  Me.DBGTransacciones.Columns("DescripcionMovimiento").Caption = "Descripcion" '3
  Me.DBGTransacciones.Columns("FacturaNo").Width = 1000 '4
  Me.DBGTransacciones.Columns("FacturaNo").Width = 1000 '4
  Me.DBGTransacciones.Columns("FacturaNo").Button = True '4
  Me.DBGTransacciones.Columns("ChequeNo").Width = 1000
  Me.DBGTransacciones.Columns("ChequeNo").Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns("Clave").Button = True
  Me.DBGTransacciones.Columns("Clave").Locked = True
  Me.DBGTransacciones.Columns("Clave").Width = 800
  Me.DBGTransacciones.Columns("TCambio").Caption = "Tasa Cambio"
  Me.DBGTransacciones.Columns("TCambio").Locked = True
  Me.DBGTransacciones.Columns("TCambio").NumberFormat = "##,##0.000000"
  Me.DBGTransacciones.Columns("TCambio").Width = 1200
  Me.DBGTransacciones.Columns("TCambio").Locked = True
  Me.DBGTransacciones.Columns("Debito").Width = 1200
  Me.DBGTransacciones.Columns("Debito").NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns("Credito").Width = 1200
  Me.DBGTransacciones.Columns("Credito").NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns("FechaTransaccion").Visible = False
  Me.DBGTransacciones.Columns("NPeriodo").Visible = False
  Me.DBGTransacciones.Columns("NTransaccion").Visible = False
  Me.DBGTransacciones.Columns("Fuente").Visible = False
  Me.DBGTransacciones.Columns("FechaTasas").Visible = False
  Me.DBGTransacciones.Columns("NumeroMovimiento").Visible = False
  Me.DBGTransacciones.Columns("Periodo").Visible = False
  Me.DBGTransacciones.Columns("FechaDescuento").Visible = False
  Me.DBGTransacciones.Columns("DescuentoDisponible").Visible = False
  Me.DBGTransacciones.Columns("FechaVence").Visible = False
  Me.DBGTransacciones.Columns("CodCuentaProveedor").Visible = False
  Me.DBGTransacciones.Columns("TipoFactura").Visible = False
  FrmTransacciones.DBGTransacciones.Columns("NTransaccion").Visible = False
 'columna tasa de cambio
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub Text2_Change()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo TipoErrs
  If Not Val(Me.TxtDiferencia.Text) = 0 Then
   MsgBox "El Movimiento esta Desbalanceado", vbCritical, "Sistema Contable"
   Cancel = 1
  End If
  If Salir = False Then
    Cancel = 1
  End If
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub List1_Click()
Me.DBGTransacciones.Columns(6).Text = Me.List1.Text
If QUIEN <> "Grid" Then
    Select Case List1.Text
      Case "Debito"
            Me.DBGTransacciones.PostMsg (3)
         
      Case "Credito"
          Me.DBGTransacciones.PostMsg (4)
      
    End Select
 End If
 
 '////////Verifico la clave del movimiento//////////
 Clave = Me.DBGTransacciones.Columns(6).Text
     If Not ClaveAnt = Clave Then
       If ClaveAnt = "Debito" Then
         Debito = Val(Me.DBGTransacciones.Columns(8).Text)
         TotalDebito = TotalDebito - Debito
         Me.TxtDebito.Text = Format(TotalDebito, "##,##0.00")
         TotalDiferencia = TotalDebito - TotalCredito
         Me.TxtDiferencia.Text = Format(TotalDiferencia, "##,##0.00")
         Me.DBGTransacciones.Columns(8).Text = "0.00"

       ElseIf ClaveAnt = "Credito" Then
         Credito = Val(Me.DBGTransacciones.Columns(9).Text)
         TotalCredito = TotalCredito - Credito
         Me.TxtCredito.Text = Format(TotalCredito, "##,##0.00")
         TotalDiferencia = TotalDebito - TotalCredito
         Me.TxtDiferencia.Text = Format(TotalDiferencia, "##,##0.00")
        Me.DBGTransacciones.Columns(9).Text = "0.00"

       End If
     End If
 
 
List1.Visible = False
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  Me.DBGTransacciones.Columns(6).Text = Me.List1.Text
   Select Case List1.Text
   Case "Debito"
         Me.DBGTransacciones.PostMsg (3)
      
   Case "Credito"
       Me.DBGTransacciones.PostMsg (4)
   
 End Select
 
  '////////Verifico la clave del movimiento//////////
 Clave = Me.DBGTransacciones.Columns(6).Text
     If Not ClaveAnt = Clave Then
       If ClaveAnt = "Debito" Then
         Me.DBGTransacciones.Columns(8).Text = "0.00"
       ElseIf ClaveAnt = "Credito" Then
        Me.DBGTransacciones.Columns(9).Text = "0.00"
       
       End If
     End If
   List1.Visible = False
 End If
End Sub

Private Sub List1_LostFocus()
Me.DBGTransacciones.Columns(6).Text = Me.List1.Text
' Select Case List1.Text
'   Case "Debito"
'         Me.DBGTransacciones.PostMsg (3)
'
'   Case "Credito"
'       Me.DBGTransacciones.PostMsg (4)
'
' End Select
'////////Verifico la clave del movimiento//////////
 Clave = Me.DBGTransacciones.Columns(6).Text
     If Not ClaveAnt = Clave Then
       If ClaveAnt = "Debito" Then
         Me.DBGTransacciones.Columns(8).Text = "0.00"
       ElseIf ClaveAnt = "Credito" Then
        Me.DBGTransacciones.Columns(9).Text = "0.00"
       
       End If
     End If
List1.Visible = False
End Sub

Private Sub RadioButton1_Click()

End Sub

Private Sub OptFacturaCompra_Click()
 If Me.OptFacturaCompra.Value = True Then
    Me.LblProveedor.Caption = "Proveedor"
    Me.AdoProveedores.RecordSource = "SELECT  CodCuentas, DescripcionCuentas, TipoCuenta From Cuentas WHERE (TipoCuenta = 'Cuentas x Pagar')    "
    Me.AdoProveedores.Refresh
    Me.LblNombres.Caption = ""
    Me.TDBProveedor.Text = ""
 ElseIf Me.OptFacturaVenta.Value = True Then
    Me.LblProveedor.Caption = "Cliente"
    Me.AdoProveedores.RecordSource = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta From Cuentas WHERE     (TipoCuenta = 'Cuentas x Cobrar')"
    Me.AdoProveedores.Refresh
    Me.LblNombres.Caption = ""
    Me.TDBProveedor.Text = ""
 End If
End Sub

Private Sub OptFacturaVenta_Click()
 If Me.OptFacturaCompra.Value = True Then
    Me.LblProveedor.Caption = "Proveedor"
    Me.AdoProveedores.RecordSource = "SELECT  CodCuentas, DescripcionCuentas, TipoCuenta From Cuentas WHERE (TipoCuenta = 'Cuentas x Pagar')    "
    Me.AdoProveedores.Refresh
    Me.LblNombres.Caption = ""
    Me.TDBProveedor.Text = ""
 ElseIf Me.OptFacturaVenta.Value = True Then
    Me.LblProveedor.Caption = "Cliente"
    Me.AdoProveedores.RecordSource = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta From Cuentas WHERE     (TipoCuenta = 'Cuentas x Cobrar')"
    Me.AdoProveedores.Refresh
    Me.LblNombres.Caption = ""
    Me.TDBProveedor.Text = ""
 End If
End Sub

Private Sub SmartButton1_Click()


On Error GoTo TipoErrs
  Dim Respuesta, Rsp, CodigoCuenta As String, FechaTransaccion As String, NumeroTransaccion As Double
  Dim NTransaccion As Double
  
'  If Not Me.DBGTransacciones.Columns(8).Text = "0.00" Then
'    MsgBox "Debe llenar de Cero el campo del Debito"
'    Exit Sub
'  End If
'
'   If Not Me.DBGTransacciones.Columns(9).Text = "0.00" Then
'    MsgBox "Debe llenar de Cero el campo del Credito"
'    Exit Sub
'  End If

   NTransaccion = Me.DBGTransacciones.Columns(12).Text
  
  Set Rsp = DtaTransacciones.Recordset
  Respuesta = MsgBox("Esta seguro de Borrar la Linea?", vbYesNo, "Transaccion No.: " & Me.TxtNTransacciones.Text)
   If Respuesta = 6 Then
     If Me.DBGTransacciones.Columns(6).Text = "Debito" Then
       Debito = Me.DBGTransacciones.Columns(8).Text
       TotalDebito = TotalDebito - Debito
       Me.TxtDebito.Text = Format(TotalDebito, "##,##0.00")
       TotalDiferencia = TotalDebito - TotalCredito
       Me.TxtDiferencia.Text = Format(TotalDiferencia, "##,##0.00")
     Else
       Credito = Me.DBGTransacciones.Columns(9).Text
       TotalCredito = TotalCredito - Credito
       Me.TxtCredito.Text = Format(TotalCredito, "##,##0.00")
       TotalDiferencia = TotalDebito - TotalCredito
       Me.TxtDiferencia.Text = Format(TotalDiferencia, "##,##0.00")
     End If
   
   If Not Me.DBGTransacciones.Columns(0).Text = "" Then
    CodigoCuenta = Me.DBGTransacciones.Columns(0).Text
    FechaTransaccion = Format(Me.TxtFecha.Value, "YYYY-MM-DD")
    NumeroTransaccion = Me.TxtNTransacciones.Text
    Me.AdoBuscar.RecordSource = "SELECT  * From Transacciones WHERE (CodCuentas = '" & CodigoCuenta & "') AND (FechaTransaccion = CONVERT(DATETIME, '" & FechaTransaccion & "', 102)) AND (NumeroMovimiento = " & NumeroTransaccion & ") AND (NTransaccion = " & NTransaccion & ")"
    Me.AdoBuscar.Refresh
    If Not Me.AdoBuscar.Recordset.EOF Then
     Me.AdoBuscar.Recordset.Delete
    
    End If
    
    
   End If
  End If
  
  
  If Not CodigoUsuario = 0 Then
Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Transacciones'))"
Me.DtaNacceso.Refresh
If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdGrabar.Enabled = False
   Me.TxtFecha.Enabled = False
   Me.Frame1.Enabled = False
   Me.DBGTransacciones.Enabled = False
   Me.CmdBuscarEmpleado.Enabled = False
      Me.TxtFecha.Enabled = False
Else
  Me.DBGTransacciones.Enabled = True
End If
Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Transacciones'))"
Me.DtaNacceso.Refresh
If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdBorrar.Enabled = False
   Me.SmartButton1.Enabled = False
End If
End If

  DtaTransacciones.Refresh
  Me.DBGTransacciones.Columns(0).Button = True
  Me.DBGTransacciones.Columns(1).Locked = True
  Me.DBGTransacciones.Columns(1).Locked = True
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Button = True
  Me.DBGTransacciones.Columns(6).Locked = True
  Me.DBGTransacciones.Columns(0).Width = 1500
  Me.DBGTransacciones.Columns(2).Width = 1000
  Me.DBGTransacciones.Columns(3).Caption = "Descripcion"
  Me.DBGTransacciones.Columns(4).Width = 1000
  Me.DBGTransacciones.Columns(4).Button = True
  Me.DBGTransacciones.Columns(5).Width = 1000
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Width = 800
  Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
  Me.DBGTransacciones.Columns(7).Locked = True
  Me.DBGTransacciones.Columns(7).NumberFormat = "##,##0.000000"
  Me.DBGTransacciones.Columns(7).Width = 1200
  Me.DBGTransacciones.Columns(8).Width = 1200
  Me.DBGTransacciones.Columns(8).NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns(9).Width = 1200
  Me.DBGTransacciones.Columns(9).NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns(10).Visible = False
  Me.DBGTransacciones.Columns(11).Visible = False
  Me.DBGTransacciones.Columns(12).Visible = False
  Me.DBGTransacciones.Columns(13).Visible = False
  Me.DBGTransacciones.Columns(14).Visible = False
  Me.DBGTransacciones.Columns(15).Visible = False
  Me.DBGTransacciones.Columns(16).Visible = False
  Me.DBGTransacciones.Columns(17).Visible = False
  Me.DBGTransacciones.Columns(18).Visible = False
  Me.DBGTransacciones.Columns(19).Visible = False
  FrmTransacciones.DBGTransacciones.Columns(20).Visible = False
  FrmTransacciones.DBGTransacciones.Columns(21).Visible = False
  FrmTransacciones.DBGTransacciones.Columns(22).Visible = False
  Me.DBGTransacciones.Columns(7).Locked = True 'columna tasa de cambio
  
 Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub TDBGridFechas_LostFocus()
' Me.TDBGridFechas.Visible = False
End Sub

Private Sub TDBProveedor_Change()
Me.TxtProveedor.Text = Me.TDBProveedor.Text
End Sub

Private Sub TDBProveedor_ItemChange()
Me.TxtProveedor.Text = Me.TDBProveedor.Text
End Sub

Private Sub TDBProveedor_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  Me.TxtProveedor.Text = Me.TDBProveedor.Text
 End If
End Sub

Private Sub TDBProveedor_SelChange(Cancel As Integer)
Me.TxtProveedor.Text = Me.TDBProveedor.Text
End Sub

Private Sub TxtFecha_GotFocus()
On Error GoTo TipoErrs
Dim Fechas1 As String, Fechas2 As String
 Me.DBGTransacciones.Enabled = True
 mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 Fechas1 = Format(FechaIni, "yyyy/mm/dd")
 Fechas2 = Format(FechaFin, "yyyy/mm/dd")
 
 Me.DtaConsulta.RecordSource = "SELECT NPeriodo, NumeroTabla, FechaPeriodo, EstadoPeriodo, NTransacciones, Periodo From Periodos WHERE     (FechaPeriodo BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas2 & "', 102))"
 ' Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  Me.TxtPeriodo.Text = DtaConsulta.Recordset("Periodo")
  NumeroPeriodo = DtaConsulta.Recordset("NPeriodo")
  NumeroTransaccion = DtaConsulta.Recordset("NTransacciones")
  EstadoPeriodo = DtaConsulta.Recordset("EstadoPeriodo")
  If EstadoPeriodo = "B" Then
   MsgBox "El Periodo Esta Bloqueado", vbCritical, "Sistema Contable"
   Me.TxtFecha.SetFocus
   Me.DBGTransacciones.Enabled = False
   TxtFecha.Enabled = True
   Me.TxtPeriodo.Enabled = True
   Me.TxtFuente.Enabled = True
   Me.TxtNTransacciones.Enabled = True
   Exit Sub
  ElseIf EstadoPeriodo = "C" Then
  MsgBox "El Periodo esta Cerrado", vbCritical, "Sistema Contable"
  Me.TxtFecha.SetFocus
  TxtFecha.Enabled = True
  Me.TxtPeriodo.Enabled = True
  Me.TxtFuente.Enabled = True
  Me.TxtNTransacciones.Enabled = True
  Me.DBGTransacciones.Enabled = False
  Exit Sub
  Else
   Me.DBGTransacciones.Enabled = True
  If Not CodigoUsuario = 0 Then
   Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Transacciones'))"
   Me.DtaNacceso.Refresh
   If Me.DtaNacceso.Recordset.EOF Then
    Me.DBGTransacciones.Enabled = False
   Else
     Me.DBGTransacciones.Enabled = True
   End If
      
  End If
  End If
 Else
   MsgBox "La Fecha esta fuera del Rango de Periodos", vbCritical, "Sistema Contable"
   Me.DBGTransacciones.Enabled = False
   TxtFecha.Enabled = True
   Me.TxtPeriodo.Enabled = True
   Me.TxtFuente.Enabled = True
   Me.TxtNTransacciones.Enabled = True
   Exit Sub
 End If
 
 '///////Verifico si esta registrada la fecha de la tasa//////

NumFecha = Me.TxtFecha.Value
Fechas = Format(Me.TxtFecha.Value, "yyyy/mm/dd")
Me.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas & "', 102))ORDER BY FechaTasas"
'DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) = " & NumFecha & "))ORDER BY Tasas.FechaTasas"
DtaTasas.Refresh

If Not DtaTasas.Recordset.EOF Then
Fecha = Format(DtaTasas.Recordset("FechaTasas"), "dd/mm/yyyy")
   
    Encontrado = True
    Cambio = DtaTasas.Recordset("MontoCordobas")
    MDIPrimero.StatusBar2.Panels(2) = "Tasa Dolar: " & Format(Cambio, "##,##0.00") & "    " & "Tasa Libras: " & Format(DtaTasas.Recordset("MontoLibras"), "##,##0.00")
End If

If Not Encontrado Then
  MsgBox "La tasa de esta Fecha no ha sido Grabada"
  Cancel = 100
  Tasa = False
  frmTasa2.Show 1
End If

Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub TxtFecha_LostFocus()
Dim NumFecha As Long, Fechas As String
mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  Me.TxtPeriodo.Text = DtaConsulta.Recordset("Periodo")
 
 End If
 
 '///////Verifico si esta registrada la fecha de la tasa//////

NumFecha = Me.TxtFecha.Value
Fechas = Format(Me.TxtFecha.Value, "yyyy/mm/dd")
Me.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas & "', 102))ORDER BY FechaTasas"
'DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) = " & NumFecha & "))ORDER BY Tasas.FechaTasas"
DtaTasas.Refresh

If Not DtaTasas.Recordset.EOF Then
Fecha = Format(DtaTasas.Recordset("FechaTasas"), "dd/mm/yyyy")
   
    Encontrado = True
    Cambio = DtaTasas.Recordset("MontoCordobas")
    MDIPrimero.StatusBar2.Panels(2) = "Tasa Dolar: " & Format(Cambio, "##,##0.00") & "    " & "Tasa Libras: " & Format(DtaTasas.Recordset("MontoLibras"), "##,##0.00")
End If

If Not Encontrado Then
  MsgBox "La tasa de esta Fecha no ha sido Grabada"
  Tasa = False
'  frmTasa2.Show 1
End If
 
End Sub



Private Sub TxtFuente_GotFocus()
On Error GoTo TipoErrs
 Me.TxtFecha.Enabled = False
 Me.TxtPeriodo.Enabled = False
 Me.TxtNTransacciones.Enabled = False
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub TxtFuente_LostFocus()
mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  Me.TxtPeriodo.Text = DtaConsulta.Recordset("Periodo")
  NumeroPeriodo = DtaConsulta.Recordset("NPeriodo")
  NumeroTransaccion = DtaConsulta.Recordset("NTransacciones")
  EstadoPeriodo = DtaConsulta.Recordset("EstadoPeriodo")
  If EstadoPeriodo = "B" Then
   MsgBox "El Periodo Esta Bloqueado", vbCritical, "Sistema Contable"
   'Me.TxtFecha.SetFocus
   Me.DBGTransacciones.Enabled = False
   TxtFecha.Enabled = True
   Me.TxtPeriodo.Enabled = True
   Me.TxtFuente.Enabled = True
   Me.TxtNTransacciones.Enabled = True
   Exit Sub
  ElseIf EstadoPeriodo = "C" Then
  MsgBox "El Periodo esta Cerrado", vbCritical, "Sistema Contable"
'  Me.TxtFecha.SetFocus
  Me.DBGTransacciones.Enabled = False
  TxtFecha.Enabled = True
  Me.TxtPeriodo.Enabled = True
  Me.TxtFuente.Enabled = True
  Me.TxtNTransacciones.Enabled = True
  Exit Sub
  Else
   Me.DBGTransacciones.Enabled = True
  End If
 Else
   MsgBox "La Fecha esta fuera del Rango de Periodos", vbCritical, "Sistema Contable"
   Me.DBGTransacciones.Enabled = False
   TxtFecha.Enabled = True
   Me.TxtPeriodo.Enabled = True
   Me.TxtFuente.Enabled = True
   Me.TxtNTransacciones.Enabled = True
   Exit Sub
 End If
 
 '///////Verifico si esta registrada la fecha de la tasa//////

NumFecha = Me.TxtFecha.Value
DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) = " & NumFecha & "))ORDER BY Tasas.FechaTasas"
DtaTasas.Refresh

If Not DtaTasas.Recordset.EOF Then
Fecha = Format(DtaTasas.Recordset("FechaTasas"), "dd/mm/yyyy")
   
    Encontrado = True
    Cambio = DtaTasas.Recordset("MontoCordobas")
    MDIPrimero.StatusBar2.Panels(2) = "Tasa Dolar: " & Format(Cambio, "##,##0.00") & "    " & "Tasa Libras: " & Format(DtaTasas.Recordset("MontoLibras"), "##,##0.00")
End If

If Not Encontrado Then
  MsgBox "La tasa de esta Fecha no ha sido Grabada"
  Cancel = 100
  Tasa = False
  frmTasa2.Show 1
End If



 Me.TxtFecha.Enabled = False
 Me.TxtPeriodo.Enabled = False
 Me.TxtFuente.Enabled = False
 Me.TxtNTransacciones.Enabled = False
 
End Sub

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    DTPFechaVence.SetFocus
 
 End If
End Sub

Private Sub TxtNTransacciones_LostFocus()
On Error GoTo TipoErrs
 mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  Me.TxtPeriodo.Text = DtaConsulta.Recordset("Periodo")
  NumeroPeriodo = DtaConsulta.Recordset("NPeriodo")
  NumeroTransaccion = DtaConsulta.Recordset("NTransacciones")
  EstadoPeriodo = DtaConsulta.Recordset("EstadoPeriodo")
  If EstadoPeriodo = "B" Then
   MsgBox "El Periodo Esta Bloqueado", vbCritical, "Sistema Contable"
   Me.TxtFecha.SetFocus
   Me.DBGTransacciones.Enabled = False
   TxtFecha.Enabled = True
   Me.TxtPeriodo.Enabled = True
   Me.TxtFuente.Enabled = True
   Me.TxtNTransacciones.Enabled = True
   Exit Sub
  ElseIf EstadoPeriodo = "C" Then
  MsgBox "El Periodo esta Cerrado", vbCritical, "Sistema Contable"
  Me.TxtFecha.SetFocus
  Me.DBGTransacciones.Enabled = False
  TxtFecha.Enabled = True
  Me.TxtPeriodo.Enabled = True
  Me.TxtFuente.Enabled = True
  Me.TxtNTransacciones.Enabled = True
  Exit Sub
  Else
   Me.DBGTransacciones.Enabled = True
  End If
 Else
   MsgBox "La Fecha esta fuera del Rango de Periodos", vbCritical, "Sistema Contable"
   Me.DBGTransacciones.Enabled = False
   TxtFecha.Enabled = True
   Me.TxtPeriodo.Enabled = True
   Me.TxtFuente.Enabled = True
   Me.TxtNTransacciones.Enabled = True
   Exit Sub
 End If
 
 '///////Verifico si esta registrada la fecha de la tasa//////

NumFecha = Me.TxtFecha.Value
DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) = " & NumFecha & "))ORDER BY Tasas.FechaTasas"
DtaTasas.Refresh

If Not DtaTasas.Recordset.EOF Then
Fecha = Format(DtaTasas.Recordset("FechaTasas"), "dd/mm/yyyy")
   
    Encontrado = True
    Cambio = DtaTasas.Recordset("MontoCordobas")
    MDIPrimero.StatusBar2.Panels(2) = "Tasa Dolar: " & Format(Cambio, "##,##0.00") & "    " & "Tasa Libras: " & Format(DtaTasas.Recordset("MontoLibras"), "##,##0.00")
End If

If Not Encontrado Then
  MsgBox "La tasa de esta Fecha no ha sido Grabada"
  Cancel = 100
  Tasa = False
  frmTasa2.Show 1
End If

'///////////////////Bloqueo los datos ////
 Me.TxtFecha.Enabled = False
 Me.TxtNTransacciones.Enabled = False

'//////////////////Agrego una nueva Transaccion///////////////
 Exit Sub
TipoErrs:
 ControlErrores
 
End Sub

Private Sub TxtProveedor_Change()
  Me.AdoBuscar.RecordSource = "SELECT * From Cuentas WHERE (CodCuentas = '" & Me.TxtProveedor.Text & "')"
  Me.AdoBuscar.Refresh
  If Not Me.AdoBuscar.Recordset.EOF Then
        If Not IsNull(Me.AdoBuscar.Recordset("CausaIva")) Then
            If AdoBuscar.Recordset("CausaIva") = True Then
              Me.OptIva.Value = True
            Else
             Me.OptIva.Value = False
            End If
         End If
         
         If Not IsNull(AdoBuscar.Recordset("CausaRetencion")) Then
            If AdoBuscar.Recordset("CausaRetencion") = True Then
              Me.OptRetencion.Value = True
            Else
              Me.OptRetencion.Value = False
            End If
         End If
     
      Me.LblNombres.Caption = Me.AdoBuscar.Recordset("DescripcionCuentas")
  End If
End Sub
