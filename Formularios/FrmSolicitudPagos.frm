VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmSolicitudPagos 
   Caption         =   "Solicitud de Pagos"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14475
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   14475
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Chk15 
      BackColor       =   &H00F7E7DE&
      Caption         =   "IVA 15%"
      Height          =   375
      Left            =   12720
      TabIndex        =   91
      Top             =   3840
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox Chk3 
      BackColor       =   &H00F7E7DE&
      Caption         =   "Retencion 3%"
      Height          =   375
      Left            =   12720
      TabIndex        =   90
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CheckBox Chk10 
      BackColor       =   &H00F7E7DE&
      Caption         =   "Retencion 10%"
      Height          =   375
      Left            =   12720
      TabIndex        =   80
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CheckBox Chk7 
      BackColor       =   &H00F7E7DE&
      Caption         =   "Retencion 7%"
      Height          =   375
      Left            =   12720
      TabIndex        =   79
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CheckBox Chk2 
      BackColor       =   &H00F7E7DE&
      Caption         =   "Retencion 2%"
      Height          =   375
      Left            =   12720
      TabIndex        =   78
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CheckBox Chk1 
      BackColor       =   &H00F7E7DE&
      Caption         =   "Retencion 1%"
      Height          =   375
      Left            =   12720
      TabIndex        =   77
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   14760
      TabIndex        =   70
      Top             =   2280
      Width           =   11415
      Begin VB.TextBox TxtNTransacciones 
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   "0"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox TxtFuente 
         Height          =   285
         Left            =   10320
         TabIndex        =   72
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtPeriodo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Fuente"
         Height          =   255
         Left            =   9600
         TabIndex        =   76
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Transaccion No."
         Height          =   255
         Left            =   5640
         TabIndex        =   75
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Periodo"
         Height          =   255
         Left            =   3000
         TabIndex        =   74
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   810
      ItemData        =   "FrmSolicitudPagos.frx":0000
      Left            =   3600
      List            =   "FrmSolicitudPagos.frx":000A
      TabIndex        =   51
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
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
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "0.00"
      Top             =   6960
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
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "0.00"
      Top             =   6960
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
      Left            =   12960
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "0.00"
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   3840
      TabIndex        =   47
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   5040
      TabIndex        =   46
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   7800
      TabIndex        =   45
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton SmartButton1 
      Caption         =   "Borrar Linea"
      Height          =   375
      Left            =   6240
      TabIndex        =   44
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   13200
      TabIndex        =   43
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton CmdMemoriza 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "Anterior"
      Height          =   375
      Left            =   1320
      TabIndex        =   41
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton CmdSiguiente 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   2520
      TabIndex        =   40
      Top             =   7680
      Width           =   1095
   End
   Begin VB.PictureBox TDBGridFechas 
      Height          =   2055
      Left            =   4920
      ScaleHeight     =   1995
      ScaleWidth      =   6075
      TabIndex        =   18
      Top             =   4800
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox TxtProveedor 
         Height          =   375
         Left            =   1800
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   2280
         Visible         =   0   'False
         Width           =   1575
      End
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
         Picture         =   "FrmSolicitudPagos.frx":001F
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   4920
         TabIndex        =   21
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4920
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtMontoCheque 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin XtremeSuiteControls.CheckBox ChkFactura 
         Height          =   255
         Left            =   240
         TabIndex        =   24
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
         TabIndex        =   25
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
            TabIndex        =   26
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
            TabIndex        =   27
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
      Begin TrueOleDBList80.TDBCombo TDBProveedor 
         Bindings        =   "FrmSolicitudPagos.frx":016D
         Height          =   315
         Left            =   1200
         TabIndex        =   28
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
         _PropDict       =   $"FrmSolicitudPagos.frx":018A
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
         TabIndex        =   29
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
            TabIndex        =   30
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
            TabIndex        =   31
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
         TabIndex        =   32
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   75759617
         CurrentDate     =   38918
      End
      Begin MSComCtl2.DTPicker DTPFechaVence 
         Height          =   300
         Left            =   3240
         TabIndex        =   33
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   75759617
         CurrentDate     =   38918
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
      Begin VB.Label Label16 
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
         TabIndex        =   38
         Top             =   550
         Width           =   1455
      End
      Begin VB.Label Label17 
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
         TabIndex        =   37
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label18 
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
         TabIndex        =   36
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Monto"
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
         Left            =   1680
         TabIndex        =   35
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label LblNombres 
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1680
         Width           =   4455
      End
   End
   Begin VB.CheckBox ChkVentana 
      Caption         =   "Mostrar Vtana Factura"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   6960
      Width           =   2295
   End
   Begin TrueOleDBList80.TDBCombo DBCodigo 
      Bindings        =   "FrmSolicitudPagos.frx":0234
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      _DropdownWidth  =   8811
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
      DropdownPosition=   1
      Locked          =   0   'False
      ScrollTrack     =   0   'False
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      AddItemSeparator=   ";"
      _PropDict       =   $"FrmSolicitudPagos.frx":024C
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
   Begin XtremeSuiteControls.GroupBox GroupBox5 
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
      _Version        =   786432
      _ExtentX        =   2990
      _ExtentY        =   7435
      _StockProps     =   79
      Caption         =   "Informacion"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox TxtCodigoEmpleado 
         Height          =   375
         Left            =   240
         TabIndex        =   93
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox ChkCheque 
         BackColor       =   &H00F7E7DE&
         Caption         =   "Imprimir Egreso, Dolares"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox TxtSaldoActual 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   1200
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmSolicitudPagos.frx":02F6
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmSolicitudPagos.frx":036C
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTasa 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmSolicitudPagos.frx":03E6
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox4 
      Height          =   3015
      Left            =   1920
      TabIndex        =   7
      Top             =   1320
      Width           =   10695
      _Version        =   786432
      _ExtentX        =   18865
      _ExtentY        =   5318
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   240
         TabIndex        =   87
         Top             =   840
         Width           =   6255
         Begin VB.OptionButton Option2 
            Caption         =   "Cancelacion"
            Height          =   255
            Left            =   2040
            TabIndex        =   89
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Anticipo"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.TextBox TxtMemo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   645
         Left            =   240
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   2280
         Width           =   10215
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
         Left            =   6600
         Picture         =   "FrmSolicitudPagos.frx":0444
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox TxtLetras 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1680
         Width           =   10215
      End
      Begin VB.TextBox TxtNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   6255
      End
      Begin MSMask.MaskEdBox TxtMonto 
         Height          =   255
         Left            =   8640
         TabIndex        =   12
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16761024
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmSolicitudPagos.frx":0592
         TabIndex        =   13
         Top             =   240
         Width           =   3975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   8040
         OleObjectBlob   =   "FrmSolicitudPagos.frx":0620
         TabIndex        =   14
         Top             =   1320
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmSolicitudPagos.frx":0688
         TabIndex        =   15
         Top             =   2040
         Width           =   1575
      End
      Begin MSMask.MaskEdBox TxtRetenciones 
         Height          =   255
         Left            =   8640
         TabIndex        =   81
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16761024
         Enabled         =   0   'False
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   7560
         OleObjectBlob   =   "FrmSolicitudPagos.frx":0704
         TabIndex        =   82
         Top             =   960
         Width           =   975
      End
      Begin MSMask.MaskEdBox TxtIVa 
         Height          =   255
         Left            =   8640
         TabIndex        =   83
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16761024
         Enabled         =   0   'False
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   8160
         OleObjectBlob   =   "FrmSolicitudPagos.frx":0778
         TabIndex        =   84
         Top             =   600
         Width           =   375
      End
      Begin MSMask.MaskEdBox TxtSubTotal 
         Height          =   255
         Left            =   8640
         TabIndex        =   85
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16761024
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   255
         Left            =   7680
         OleObjectBlob   =   "FrmSolicitudPagos.frx":07DC
         TabIndex        =   86
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LblSimbolo 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto C$"
         Height          =   255
         Left            =   5040
         TabIndex        =   16
         Top             =   480
         Width           =   855
      End
   End
   Begin TrueOleDBGrid80.TDBGrid DBGTransacciones 
      Bindings        =   "FrmSolicitudPagos.frx":084C
      Height          =   2415
      Left            =   120
      TabIndex        =   52
      Top             =   4440
      Width           =   14175
      _ExtentX        =   25003
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
      Splits(0).Caption=   "Movimento de Egresos"
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
      AllowAddNew     =   -1  'True
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=8,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HF7C1C5&,.fgcolor=&H0&,.bold=-1"
      _StyleDefs(20)  =   ":id=22,.fontsize=1200,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(21)  =   ":id=22,.fontname=Script MT Bold"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Named:id=33:Normal"
      _StyleDefs(41)  =   ":id=33,.parent=0"
      _StyleDefs(42)  =   "Named:id=34:Heading"
      _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(44)  =   ":id=34,.wraptext=-1"
      _StyleDefs(45)  =   "Named:id=35:Footing"
      _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   "Named:id=36:Selected"
      _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(49)  =   "Named:id=37:Caption"
      _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(51)  =   "Named:id=38:HighlightRow"
      _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=39:EvenRow"
      _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(55)  =   "Named:id=40:OddRow"
      _StyleDefs(56)  =   ":id=40,.parent=33"
      _StyleDefs(57)  =   "Named:id=41:RecordSelector"
      _StyleDefs(58)  =   ":id=41,.parent=34"
      _StyleDefs(59)  =   "Named:id=42:FilterBar"
      _StyleDefs(60)  =   ":id=42,.parent=33"
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   1095
      Left            =   1920
      TabIndex        =   53
      Top             =   120
      Width           =   12375
      _Version        =   786432
      _ExtentX        =   21828
      _ExtentY        =   1931
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin VB.CommandButton CmdConsultaBanco 
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
         Left            =   3480
         Picture         =   "FrmSolicitudPagos.frx":086B
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox CmbMoneda 
         Height          =   315
         ItemData        =   "FrmSolicitudPagos.frx":09B9
         Left            =   3600
         List            =   "FrmSolicitudPagos.frx":09C6
         TabIndex        =   55
         Top             =   680
         Width           =   1815
      End
      Begin VB.TextBox TxtNombreBanco 
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   240
         Width           =   6495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   2760
         OleObjectBlob   =   "FrmSolicitudPagos.frx":09E5
         TabIndex        =   56
         Top             =   720
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmSolicitudPagos.frx":0A4F
         TabIndex        =   57
         Top             =   1200
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmSolicitudPagos.frx":0ACF
         TabIndex        =   58
         Top             =   700
         Width           =   495
      End
      Begin MSDataListLib.DataCombo DBEmpleado 
         Bindings        =   "FrmSolicitudPagos.frx":0B37
         Height          =   315
         Left            =   1680
         TabIndex        =   59
         Top             =   1200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Left            =   5040
         OleObjectBlob   =   "FrmSolicitudPagos.frx":0B52
         TabIndex        =   60
         Top             =   1200
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblSaldo 
         Height          =   375
         Left            =   7080
         OleObjectBlob   =   "FrmSolicitudPagos.frx":0BD6
         TabIndex        =   61
         Top             =   1200
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker TxtFecha 
         Height          =   285
         Left            =   720
         TabIndex        =   62
         Top             =   680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   75759617
         CurrentDate     =   38008
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   4200
         OleObjectBlob   =   "FrmSolicitudPagos.frx":0C34
         TabIndex        =   63
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmSolicitudPagos.frx":0CA8
         TabIndex        =   64
         Top             =   240
         Width           =   615
      End
      Begin MSDataListLib.DataCombo DBCodigo1 
         Bindings        =   "FrmSolicitudPagos.frx":0D12
         Height          =   315
         Left            =   5880
         TabIndex        =   65
         Top             =   600
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   0
         Left            =   7080
         TabIndex        =   66
         Top             =   240
         Width           =   1215
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
      Height          =   255
      Left            =   12960
      OleObjectBlob   =   "FrmSolicitudPagos.frx":0D2A
      TabIndex        =   67
      Top             =   7320
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
      Height          =   255
      Left            =   11520
      OleObjectBlob   =   "FrmSolicitudPagos.frx":0D96
      TabIndex        =   68
      Top             =   7320
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
      Height          =   255
      Left            =   3000
      OleObjectBlob   =   "FrmSolicitudPagos.frx":0E00
      TabIndex        =   69
      Top             =   6960
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc DtaDatosEmpresa 
      Height          =   375
      Left            =   4080
      Top             =   9600
      Width           =   3855
      _ExtentX        =   6800
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
      Caption         =   "DtaDatosEmpresa"
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
      Left            =   8400
      Top             =   9000
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
   Begin MSAdodcLib.Adodc DtaBancos 
      Height          =   375
      Left            =   11280
      Top             =   9240
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
      Caption         =   "DtaBancos"
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
   Begin MSAdodcLib.Adodc DtaContratista 
      Height          =   375
      Left            =   1200
      Top             =   10800
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
      Caption         =   "DtaContratista"
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
   Begin MSAdodcLib.Adodc DtaHistorial 
      Height          =   330
      Left            =   2880
      Top             =   10800
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
      Left            =   2760
      Top             =   10320
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
   Begin MSAdodcLib.Adodc DtaTransacciones 
      Height          =   375
      Left            =   6360
      Top             =   9000
      Visible         =   0   'False
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
      CommandType     =   1
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
      Left            =   8760
      Top             =   10080
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
   Begin MSAdodcLib.Adodc DtaConsulta 
      Height          =   375
      Left            =   5760
      Top             =   10560
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
      Left            =   5760
      Top             =   10080
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
      Left            =   7800
      Top             =   9360
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
   Begin MSAdodcLib.Adodc AdoCordenadas 
      Height          =   375
      Left            =   1200
      Top             =   9360
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
      Caption         =   "AdoCordenadas"
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
   Begin MSAdodcLib.Adodc AdoPendientes 
      Height          =   330
      Left            =   2520
      Top             =   9840
      Width           =   3015
      _ExtentX        =   5318
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
      Caption         =   "AdoPendientes"
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
      Left            =   720
      Top             =   9960
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
   Begin MSAdodcLib.Adodc DtaIndice 
      Height          =   375
      Left            =   2880
      Top             =   8880
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
      Caption         =   "DtaContratista"
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
Attribute VB_Name = "FrmSolicitudPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ChequeGrabado As Boolean
Private ew As cls_NumEnglishWord
Private sw As cls_NumSpanishWord

Private Sub Chk1_Click()

    Me.TxtMonto.Text = CalcularMonto

End Sub

Private Sub Chk10_Click()
    Me.TxtMonto.Text = CalcularMonto
End Sub

Private Sub Chk15_Click()
Me.TxtMonto.Text = CalcularMonto
End Sub

Private Sub Chk2_Click()
    Me.TxtMonto.Text = CalcularMonto
End Sub

Private Sub Chk3_Click()
    Me.TxtMonto.Text = CalcularMonto
End Sub

Private Sub Chk7_Click()
    Me.TxtMonto.Text = CalcularMonto
End Sub

Private Sub CmdBuscarEmpleado_Click()
On Error GoTo TipoErrs
 QueProducto = "ContratistaCheque"
 FrmConsulta.Show 1
Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub CmdConsultaBanco_Click()
On Error GoTo TipoErrs
 QueProducto = "ChequeBanco"
 FrmConsulta.Show 1
 
 Me.DBCodigo.Text = FrmConsulta.Codigo
 Me.TxtCodigoEmpleado.Text = FrmConsulta.Codigo
 
Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub DBCodigo_ItemChange()
On Error GoTo TipoErrs
Dim MontoTasa As Double, Fecha As Long
Dim SQL As String
Criterio = "CodCuentas='" & Me.DBCodigo.Text & "'"
If Me.DtaCuentas.Recordset.RecordCount > 0 Then Me.DtaCuentas.Recordset.MoveFirst
Me.DtaCuentas.Recordset.Find (Criterio)
  

If Not DtaCuentas.Recordset.EOF Then

'////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////CARGO LOS CHEQUES PENDIENTES/////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////
SQL = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, " & _
"Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, " & _
"Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, " & _
"Transacciones.NumeroMovimiento , Periodos.Periodo, Transacciones.Beneficiario " & _
"FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo " & _
"WHERE (Transacciones.ChequeNo = '#######')  AND (Transacciones.CodCuentas = '" & Me.DBCodigo.Text & "' ) AND " & _
"(Transacciones.DescripcionMovimiento <> '**********CANCELADO*************') ORDER BY Transacciones.NTransaccion"

Me.AdoPendientes.RecordSource = SQL
Me.AdoPendientes.Refresh





   Me.TxtNombreBanco.Text = Me.DtaCuentas.Recordset("DescripcionCuentas")
              TipoMoneda = Me.DtaCuentas.Recordset("TipoMoneda")
              
          Me.CmbMoneda.Text = TipoMoneda

         Select Case TipoMoneda
            Case "C�rdobas"
                      Fecha = Me.TxtFecha.Value
                      Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      Me.DtaTasas.Refresh
                If Not DtaTasas.Recordset.EOF Then
                 MontoTasa = Me.DtaTasas.Recordset("MontoCordobas")
                 If MontoTasa = 0 Then
                   MontoTasa = Format(1, "##,##0.000000")
                 End If
                 Me.LblTasa.Caption = Format((1 / MontoTasa), "##,##0.000000")
                Else
                 Me.LblTasa.Caption = Format(1, "##,##0.000000")
                End If
            
            Case "D�lares"
                    Me.LblTasa.Caption = Format(1, "##,##0.000000")
            
            Case "Libras"
                      Fecha = Me.TxtFecha.Value
                      Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      Me.DtaTasas.Refresh
                If Not DtaTasas.Recordset.EOF Then
                 MontoTasa = Me.DtaTasas.Recordset("MontoLibras")
                 Me.LblTasa.Caption = Format(MontoTasa, "##,##0.000000")
                Else
                Me.LblTasa.Caption = Format(1, "##,##0.000000")
                End If
         
         End Select



 Me.TxtFecha.Enabled = True
 If Not IsNull(Me.DtaBancos.Recordset("TipoCuenta")) Then
   TipoCuenta = Me.DtaBancos.Recordset("TipoCuenta")
 End If
 
 If TipoMoneda = "C�rdobas" Then
  Me.LblSimbolo.Caption = "Monto $"
 ElseIf TipoMoneda = "D�lares" Then
  Me.LblSimbolo.Caption = "Monto C$"
 End If
 
 
ElseIf Not Me.DBCodigo.Text = "" Then
  'MsgBox "El codigo Digitado no es correcto", vbCritical, "Sistema Contable"
  Exit Sub
Else
  Me.TxtNombreBanco.Text = ""
End If
Exit Sub
TipoErrs:
MsgBox err.Description

End Sub



Private Sub DBGTransacciones_BeforeUpdate(Cancel As Integer)
'On Error GoTo TipoErrs
 If Me.DBGTransacciones.Columns(6).Text = "" Then
   Me.DBGTransacciones.Columns(6).Text = "Debito"
 End If
 
  If Me.DBGTransacciones.Columns(8).Text = "" Then
   Me.DBGTransacciones.Columns(8).Text = 0
 End If
 
 If Me.DBGTransacciones.Columns(9).Text = "" Then
   Me.DBGTransacciones.Columns(9).Text = 0
 End If
  For i = 2 To 5
            If Me.DBGTransacciones.Columns(i).Text = "" Then Me.DBGTransacciones.Columns(i).Text = "-"
        Next i
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub DBGTransacciones_ButtonClick(ByVal ColIndex As Integer)
'On Error GoTo TipoErrs
Dim c As Variant
Select Case ColIndex
  Case 0
  QueProducto = "SolicitudPagos"
  FrmConsulta.Show 1
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
End Select

Exit Sub
TipoErrs:
 ControlErrores
End Sub


Private Sub DBGTransacciones_AfterColEdit(ByVal ColIndex As Integer)
On Error GoTo TipoErrs
Dim Descripcion As String, Cadena As String, MontoTasa As Double, Fecha As Long
Dim ClaveMovimiento As String, DescripcionMovimiento As String, SQL As String
Dim c As Variant
'Este Procedimiento es solo cuando se ejecuta directamente de Recepcion
QueProducto = "Egreso"

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
  
        Me.DBGTransacciones.Enabled = False
        
   End If
  
  End If

'/////Busco cambios en las claves del movimiento///////////


Select Case ColIndex
  Case 0
    '////////////Verifico la cuenta///////////////
       

       Me.DtaCuentas.Refresh

       Criterio = "CodCuentas='" & Me.DBGTransacciones.Columns(0).Text & " '"
       Me.DtaCuentas.Recordset.Find (Criterio)
       If Not DtaCuentas.Recordset.EOF Then
                TipoMoneda = DtaCuentas.Recordset("TipoMoneda")
                
                 Me.DBGTransacciones.Columns(1).Text = DtaCuentas.Recordset("DescripcionCuentas")
         Select Case TipoMoneda
            Case "C�rdobas"
                      Fecha = Me.TxtFecha.Value
                      Me.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE (FechaTasas = '" & Format(Me.TxtFecha.Value, "yyyymmdd") & "')"
                      Me.DtaTasas.Refresh
                If Not DtaTasas.Recordset.EOF Then
                 MontoTasa = Me.DtaTasas.Recordset("MontoCordobas")
                 If MontoTasa = 0 Then
                   MontoTasa = 1
                 End If
                 Select Case Me.CmbMoneda.Text
                  Case "C�rdobas"
                    Me.DBGTransacciones.Columns(7).Text = 1
                  Case "D�lares"
                    Me.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Libras"
                    Me.DBGTransacciones.Columns(7).Text = MontoTasa
                   ' Me.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                    
                 End Select
                Else
                
                 Me.DBGTransacciones.Columns(7).Text = 1
                End If
            
            Case "D�lares"
             Fecha = Me.TxtFecha.Value
                      Me.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE (FechaTasas = '" & Format(Me.TxtFecha.Value, "yyyymmdd") & "')"
                      Me.DtaTasas.Refresh
             If Not DtaTasas.Recordset.EOF Then
                MontoTasa = Me.DtaTasas.Recordset("MontoCordobas")
                If MontoTasa = 0 Then
                  MontoTasa = 1
                End If
            
               Select Case Me.CmbMoneda.Text
                  Case "C�rdobas"
                    Me.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                  Case "D�lares"
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
                      Me.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE (FechaTasas = '" & Format(Me.TxtFecha.Value, "yyyymmdd") & "')"
                      Me.DtaTasas.Refresh
                If Not DtaTasas.Recordset.EOF Then
                MontoTasa = Me.DtaTasas.Recordset("MontoLibras")
               Select Case Me.CmbMoneda.Text
                  Case "C�rdobas"
                    Me.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "D�lares"
                    Me.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Libras"
                    Me.DBGTransacciones.Columns(7).Text = 1

                    
                 End Select
                Else
                 Me.DBGTransacciones.Columns(7).Text = 1
                End If
         
         End Select
         
         
  Criterio = "CodCuentas='" & Me.DBCodigo.Text & "'"
  Me.DtaCuentas.Recordset.Find (Criterio)
  If Not DtaCuentas.Recordset.EOF Then
  'Criterio = "CodCuentas='" & Me.DBCodigo.Text & "'"
  'Me.DtaCuentas.Recordset.Find (Criterio)
      
    TipoCuenta = Me.DtaCuentas.Recordset("TipoCuenta")
    CodigoCuenta = Me.DtaCuentas.Recordset("CodCuentas")
   End If
  If TipoCuenta = "Bancos" Then

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
       ' Me.DtaTransacciones.Recordset.MoveLast
     
     End If
        ConsecutivoVoucher = Month(Me.TxtFecha.Value)
        If TipoCuenta = "Caja" Then
              numero = "CASH " & NumeroVoucher & "/" & ConsecutivoVoucher
        End If
        Select Case TipoMoneda
           Case "C�rdobas"
            If TipoCuenta = "Bancos" Then
              numero = "BC " & NumeroVoucher & "/" & ConsecutivoVoucher
            End If
           Case "D�lares"
            If TipoCuenta = "Bancos" Then
              numero = "BD " & NumeroVoucher & "/" & ConsecutivoVoucher
            End If
        
         End Select
        
     End If
       
       
       
         '  Cadena = Mid(Me.DBCodigo, 1, 1)
          ' Cadena = Cadena & "/" & NumeroTransaccion
           
        
   '///////////////////////////////////////////////////////////
   '//////CON ESTA CONSULTA BUSCO LA DESCRIPCION DE LA LINEA ANTERIOR//////////////////
   '/////////////////////////////////////////////////////////////////////////////////
   
            
            SQL = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta AS DescripcionCuentas, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, " & _
            "Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, " & _
            "Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, " & _
            "Transacciones.NumeroMovimiento , Periodos.Periodo " & _
            "FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo " & _
            "WHERE  (Transacciones.FechaTransaccion BETWEEN '" & Format(Me.TxtFecha.Value, "yyyymmdd") & "' And '" & Format(Me.TxtFecha.Value, "yyyymmdd") & "') AND (Transacciones.NumeroMovimiento = " & Me.TxtNTransacciones.Text & ") " & _
            "ORDER BY Transacciones.NTransaccion"
              
            Me.DtaConsulta.RecordSource = SQL
            Me.DtaConsulta.Refresh
            If Not Me.DtaConsulta.Recordset.EOF Then
              Me.DtaConsulta.Recordset.MoveLast
              If Not IsNull(Me.DtaConsulta.Recordset("DescripcionMovimiento")) Then
                 DescripcionMovimiento = Me.DtaConsulta.Recordset("DescripcionMovimiento")
              End If
              If Not IsNull(Me.DtaConsulta.Recordset("Clave")) Then
                ClaveMovimiento = Me.DtaConsulta.Recordset("Clave")
              End If
            
            End If
          

         Me.DBGTransacciones.Columns(2).Text = Cadena
         Me.DBGTransacciones.Columns(3).Text = DescripcionMovimiento
         Me.DBGTransacciones.Columns(10).Text = Format(Me.TxtFecha.Value, "dd/mm/yyyy")
         Me.DBGTransacciones.Columns(11).Text = NumeroPeriodo
         Me.DBGTransacciones.Columns(13).Text = Me.TxtFuente.Text
         Me.DBGTransacciones.Columns(14).Text = Format(Me.TxtFecha.Value, "dd/mm/yyyy")
         Me.DBGTransacciones.Columns(15).Text = Me.TxtNTransacciones.Text
         If ClaveMovimiento = "" Then
          Me.DBGTransacciones.Columns(6).Text = "Debito"
         Else
          Me.DBGTransacciones.Columns(6).Text = ClaveMovimiento
         End If
         'Me.DBGTransacciones.Columns(9).Locked = True
        
         
         
       Else
               
         MsgBox "La cuenta digitada no es correcta", vbCritical, "Sistema Contable"
         NumeroTransaccion = Me.TxtNTransacciones.Text
         'Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,  Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo,FechaDescuento,DescuentoDisponible,FechaVence FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion"
         'Me.DtaTransacciones.Refresh
  Me.DBGTransacciones.Columns(0).Button = True
    Me.DBGTransacciones.Columns(1).Locked = True
  Me.DBGTransacciones.Columns(6).Button = True
  Me.DBGTransacciones.Columns(6).Locked = True
  Me.DBGTransacciones.Columns(0).Width = 1500
  Me.DBGTransacciones.Columns(2).Width = 1000
  Me.DBGTransacciones.Columns(3).Caption = "Descripcion"
  Me.DBGTransacciones.Columns(4).Width = 1000
  Me.DBGTransacciones.Columns(4).Button = True
  Me.DBGTransacciones.Columns(5).Width = 1000
  Me.DBGTransacciones.Columns(6).Width = 800
  Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
  Me.DBGTransacciones.Columns(7).Locked = True
  Me.DBGTransacciones.Columns(7).Width = 1200
  Me.DBGTransacciones.Columns(7).NumberFormat = "##,##0.000000"
  Me.DBGTransacciones.Columns(8).Width = 1200
  Me.DBGTransacciones.Columns(8).NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns(9).Width = 1200
  Me.DBGTransacciones.Columns(9).NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns(10).Visible = False
  Me.DBGTransacciones.Columns(11).Visible = False
  Me.DBGTransacciones.Columns(12).Visible = False
  Me.DBGTransacciones.Columns(11).Visible = False
  Me.DBGTransacciones.Columns(12).Visible = False
  Me.DBGTransacciones.Columns(13).Visible = False
  Me.DBGTransacciones.Columns(14).Visible = False
  Me.DBGTransacciones.Columns(15).Visible = False
    Me.DBGTransacciones.Columns(17).Visible = False
  Me.DBGTransacciones.Columns(18).Visible = False
'  Me.DBGTransacciones.Columns(19).Visible = False
         Me.DBGTransacciones.Columns(16).Visible = False
         FrmConsulta.Show 1
         Exit Sub
       End If
     
    
 
 
       
 Case 7
 Salir = False
   '//////////Sumo los totales del Debito///////////////
    If Me.DBGTransacciones.Columns(7).Text = "" Then
      Me.DBGTransacciones.Columns(7).Text = "0.00"
    End If
    
  
    
    Debito = Me.DBGTransacciones.Columns(7).Text
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
Case 8
Salir = False
    '//////////Sumo los totales del credito///////////////
    If Me.DBGTransacciones.Columns(8).Text = "" Then
      Me.DBGTransacciones.Columns(8).Text = "0.00"
    End If
    
       
    Credito = Me.DBGTransacciones.Columns(8).Text
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
'          NumeroPeriodoAnterior = DtaPeriodos.Recordset("NPeriodo")
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
          DtaPeriodos.Recordset.MovePrevious
'          NumeroPeriodoAnterior = DtaPeriodos.Recordset("NPeriodo")
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodoAnterior & "))"
          Me.DtaHistorial.Refresh
        
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
          DtaPeriodos.Recordset.MovePrevious
'          NumeroPeriodoAnterior = DtaPeriodos.Recordset("NPeriodo")
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodoAnterior & "))"
          Me.DtaHistorial.Refresh
        
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
'On Error GoTo TipoErrs
   Select Case ColIndex
    
    Case 0
    Me.DBCodigo.Enabled = False
    Me.TxtFecha.Enabled = False
    
    
    
      mes = Month(Me.TxtFecha.Value)
      A�o = Year(Me.TxtFecha.Value)
      FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
      FechaFin = DateSerial(A�o, mes + 1, 1 - 1)
      NumFecha1 = FechaIni
      NumFecha2 = FechaFin
 
      Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
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
      Me.TxtNTransacciones.Text = NumeroTransaccion
      '////////Edito los Datos de los indices de Transacciones//////
         
          Me.DtaIndice.Recordset.AddNew
          Me.DtaIndice.Recordset("FechaTransaccion") = Format(Me.TxtFecha.Value, "dd/mm/yyyy")
          Me.DtaIndice.Recordset("NumeroMovimiento") = NumeroTransaccion
          Me.DtaIndice.Recordset("DescripcionMovimiento") = Me.DBGTransacciones.Columns(1).Text
          Me.DtaIndice.Recordset("Fuente") = Me.TxtFuente.Text
          Me.DtaIndice.Recordset("NPeriodo") = NumeroPeriodo
          If Me.CmbMoneda.Text = "D�lares" Then
            Me.DtaIndice.Recordset("TipoMoneda") = "D�lares"
          ElseIf Me.CmbMoneda.Text = "C�rdobas" Then
            Me.DtaIndice.Recordset("TipoMoneda") = "C�rdobas"
          End If
          
          
          Me.DtaIndice.Recordset.Update
      
      
      
     
     
     
     End If
   End If
   
      Case 3
      mes = Month(Me.TxtFecha.Value)
      A�o = Year(Me.TxtFecha.Value)
      FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
      FechaFin = DateSerial(A�o, mes + 1, 1 - 1)
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
        End If
        EstadoPeriodo = DtaConsulta.Recordset("EstadoPeriodo")
      End If
  Me.AdoBuscar.RecordSource = "SELECT FechaTransaccion, NumeroMovimiento, Nperiodo, DescripcionMovimiento, Fuente, TipoMoneda From IndiceTransaccion Where (NPeriodo = " & NumeroPeriodo & ") And (NumeroMovimiento = " & NumeroTransaccion & ")"
  Me.AdoBuscar.Refresh
  
   If Not Me.AdoBuscar.Recordset.EOF Then
   Me.AdoBuscar.Recordset("DescripcionMovimiento") = Me.DBGTransacciones.Columns(3).Text
   Me.AdoBuscar.Recordset.Update
   End If
   
   
  End Select
  
  Exit Sub
TipoErrs:
 ControlErrores

End Sub

Private Sub Form_Load()
Dim SQL As String

MDIPrimero.Skin1.ApplySkin hWnd
'Me.TxtFecha.Value = Format(FechaSistema, "dd/mm/yyyy")

Me.ChkVentana.BackColor = RGB(222, 231, 247)

MDIPrimero.Skin1.ApplySkin Me.CmdSiguiente.hWnd
MDIPrimero.Skin1.ApplySkin Me.CmdAnterior.hWnd
MDIPrimero.Skin1.ApplySkin Me.CmdMemoriza.hWnd
MDIPrimero.Skin1.ApplySkin Me.CmdBorrar.hWnd
MDIPrimero.Skin1.ApplySkin Me.CmdSalir.hWnd
MDIPrimero.Skin1.ApplySkin Me.CmdGrabar.hWnd
MDIPrimero.Skin1.ApplySkin Me.CmdNuevo.hWnd
MDIPrimero.Skin1.ApplySkin Me.SmartButton1.hWnd
Me.TxtFuente.Text = "EGRESO"
Primero = True
Salir = True
'On Error GoTo TipoErrs
Dim SqlCheque As String
With Me.DtaContratista
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

'With Me.DtaConsecutivo
'   .ConnectionString = Conexion
'End With

With Me.AdoBuscar
   .ConnectionString = Conexion
End With

With Me.DtaIndice
   .ConnectionString = Conexion
End With

With Me.AdoCordenadas
   .ConnectionString = Conexion
End With

'With Me.AdoProveedores
'   .ConnectionString = Conexion
'End With

With Me.DtaDatosEmpresa

   .ConnectionString = Conexion
   .RecordSource = "DatosEmpresa"
   .Refresh
End With

With Me.DtaNacceso
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Accesos"
   .Refresh
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

With Me.DtaBancos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.AdoPendientes
   .ConnectionString = Conexion
End With


    ChequeGrabado = False

Me.TxtFecha.Value = Format(Now, "dd/mm/yyyy")
Me.DBGTransacciones.Enabled = False
Me.TxtMemo.Enabled = False
Me.TxtMonto.Enabled = False
Me.TxtNombre.Enabled = False

SQL = "SELECT     Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, " & _
       "Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, " & _
       "Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, " & _
       "Transacciones.NumeroMovimiento, Periodos.Periodo, Transacciones.FechaDescuento, Transacciones.DescuentoDisponible, " & _
       "Transacciones.FechaVence,Transacciones.CodCuentaProveedor,Transacciones.TipoFactura,Transacciones.NTransaccion " & _
       "FROM  Periodos INNER JOIN " & _
       "Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo " & _
       "Where (Transacciones.NumeroMovimiento = -1) " & _
       "ORDER BY Transacciones.NTransaccion "
       
Me.DtaTransacciones.RecordSource = SQL
Me.DtaTransacciones.Refresh

Me.DtaBancos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta From Cuentas WHERE (TipoCuenta = 'Caja') OR (TipoCuenta = N'Bancos') ORDER BY Cuentas.CodCuentas"
Me.DtaBancos.Refresh
Me.DBCodigo.ListField = "CodCuentas"

 

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
  DBGTransacciones.Columns(20).Visible = False
  DBGTransacciones.Columns(21).Visible = False
  DBGTransacciones.Columns(22).Visible = False
  DBGTransacciones.Columns(7).Locked = True 'columna tasa de cambio

End Sub
  Private Sub Form_Initialize()
On Error GoTo TipoErrs
Dim SqlCheque As String
    Set ew = New cls_NumEnglishWord
    Set sw = New cls_NumSpanishWord
    'DBGdetalleCk.Columns(3).Button = True
Exit Sub
TipoErrs:
ControlErrores
End Sub
Private Sub DBGTransacciones_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
 Dim Criterio As String
 
   
 '/////////////////REVALIDO SI LA CUENTA EXISTE /////////////////////////////////
       Criterio = Me.DBGTransacciones.Columns(0).Text
       Me.AdoBuscar.RecordSource = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo From Cuentas WHERE (CodCuentas = '" & Criterio & "')"
       Me.AdoBuscar.Refresh
       If Not Me.AdoBuscar.Recordset.EOF Then
         Me.DBGTransacciones.Columns(1).Text = Me.AdoBuscar.Recordset("DescripcionCuentas")
       Else
'         MsgBox "No Existe la Cuenta", vbCritical, "Zeus Facturacion"
         Me.DBGTransacciones.Columns(0).Text = ""
         Me.DBGTransacciones.Columns(1).Text = ""
       End If
End Sub

Private Sub DBGTransacciones_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
On Error GoTo TipoErrs


If ColIndex = 8 Or ColIndex = 9 Then
 If Me.DBGTransacciones.Columns(6).Text = "Debito" Then
      'Me.DBGTransacciones.Columns(9).Locked = True
      'Me.DBGTransacciones.Columns(8).Locked = False
  ElseIf Me.DBGTransacciones.Columns(6).Text = "Credito" Then
  
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

Private Sub DBGTransacciones_AfterUpdate()
Dim MontoCheque As Double, MontoTasa As Double, Fecha As Long
Dim Fechas1 As String

 Debito = 0
 Credito = 0
 TotalDebito = 0
 TotalCredito = 0
      NumFecha1 = Me.TxtFecha.Value
      NMovimiento = Val(Me.TxtNTransacciones)
      Fechas1 = Format(Me.TxtFecha.Value, "yyyy/mm/dd")
      Me.DtaConsulta.RecordSource = "SELECT FechaTransaccion, CodCuentas, NTransaccion, NumeroMovimiento, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, TCambio, Debito,Credito From Transacciones WHERE  (FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (NumeroMovimiento = " & NMovimiento & ")"

      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
      If TxtMonto.Text = "" Then
       MontoCheque = 0
      Else
       MontoCheque = Me.TxtMonto
      End If
       Debito = Me.DtaConsulta.Recordset("Debito")
       Credito = Me.DtaConsulta.Recordset("Credito")
       TotalDebito = TotalDebito + Debito
       TotalCredito = TotalCredito + Credito
       DtaConsulta.Recordset.MoveNext
      Loop
      
       Criterio = "CodCuentas='" & Me.DBCodigo.Text & "'"
       Me.DtaCuentas.Recordset.Find (Criterio)
       If Not DtaCuentas.Recordset.EOF Then
                TipoMoneda = DtaCuentas.Recordset("TipoMoneda")

       If Not TxtMonto.Text = "" Then
        MontoCheque = Me.TxtMonto.Text
       Else
        MontoCheque = 0
       End If
         Select Case TipoMoneda
            Case "C�rdobas"
                      Fecha = Me.TxtFecha.Value
                      Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      Me.DtaTasas.Refresh
                If Not DtaTasas.Recordset.EOF Then
                 MontoTasa = Me.DtaTasas.Recordset("MontoCordobas")
                 If MontoTasa = 0 Then
                   MontoTasa = 1
                 End If
                 'MontoCheque = (1 / MontoTasa) * MontoCheque
                Else
                 'MontoCheque = 1 * MontoCheque
                End If
            
            Case "D�lares"
                  'MontoCheque = 1 * MontoCheque
            
            Case "Libras"
                      Fecha = Me.TxtFecha.Value
                      Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      Me.DtaTasas.Refresh
                If Not DtaTasas.Recordset.EOF Then
                 MontoTasa = Me.DtaTasas.Recordset("MontoLibras")
 '               MontoCheque = MontoTasa * MontoCheque
                Else
'               MontoCheque = 1 * MontoCheque
                End If
         
         End Select

End If


'Diferencia = TotalMovimientos(Me.TxtFecha.Value, Me.TxtNTransacciones)

TotalCredito = TotalCredito + MontoCheque
Me.TxtCredito.Text = Format(TotalCredito, "##,##0.00")
Me.TxtDebito.Text = Format(TotalDebito, "##,##0.00")
Me.TxtDiferencia.Text = Format(TotalDebito - TotalCredito, "##,##0.00")

Me.DBGTransacciones.SetFocus
Me.DBGTransacciones.PostMsg 2

Me.CmdNuevo.Enabled = False
End Sub

Private Sub MaskEdBox3_Change()
'On Error GoTo TipoErrs
Dim SqlContratista As String, Fecha As Long

 Debito = 0
 Credito = 0
 TotalDebito = 0
 TotalCredito = 0

If Not DBCodigo.Text = "" Then
  If TxtMonto.Text = "" Then
   Credito = 0
  Else
   Credito = Me.TxtMonto
  End If

      NumFecha1 = Me.TxtFecha.Value
      NMovimiento = Val(Me.TxtNTransacciones)
'      Me.DtaConsulta.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.CodCuentas, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, TCambio*Debito AS MDebito, TCambio*Credito AS MCredito, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito From Transacciones WHERE (((Transacciones.FechaTransaccion)=" & NumFecha1 & ") AND ((Transacciones.NumeroMovimiento)=" & NMovimiento & "))"
      Me.DtaConsulta.RecordSource = "SELECT FechaTransaccion, CodCuentas, NTransaccion, NumeroMovimiento, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, TCambio, Debito, Credito From Transacciones  " & _
                                    "WHERE (FechaTransaccion = CONVERT(DATETIME, '" & Format(Me.TxtFecha.Value, "yyyy-mm-dd") & "', 102)) AND (NumeroMovimiento = " & NMovimiento & ") AND (CodCuentas <> '" & Me.DBCodigo.Text & "')"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
       If Me.TxtMonto.Text = "" Then
        MontoCheque = 0
       Else
         MontoCheque = Me.TxtMonto
       End If
       Debito = Me.DtaConsulta.Recordset("Debito")
       Credito = Me.DtaConsulta.Recordset("Credito")
       TotalDebito = TotalDebito + Debito
       TotalCredito = TotalCredito + Credito
       DtaConsulta.Recordset.MoveNext
      Loop
      

      
       Criterio = "CodCuentas='" & Me.DBCodigo.Text & "'"
       Me.DtaCuentas.Recordset.Find (Criterio)
       If Not DtaCuentas.Recordset.EOF Then
'                TipoMoneda = DtaCuentas.Recordset("TipoMoneda")
                TipoMoneda = Me.CmbMoneda.Text
                TipoCuenta = DtaCuentas.Recordset("TipoCuenta")
      If Me.TxtMonto.Text = "" Then
         MontoCheque = 0
      Else
           MontoCheque = Me.TxtMonto.Text
      End If
     
         Select Case TipoMoneda
            Case "C�rdobas"
                      Fecha = Me.TxtFecha.Value
                      Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      Me.DtaTasas.Refresh
                If Not DtaTasas.Recordset.EOF Then
                 MontoTasa = Me.DtaTasas.Recordset("MontoCordobas")
                 If MontoTasa = 0 Then
                   MontoTasa = 1
                 End If
                 'MontoCheque = (1 / MontoTasa) * MontoCheque
                Else
                 'MontoCheque = 1 * MontoCheque
                End If
            
            Case "D�lares"
                '  MontoCheque = 1 * MontoCheque
            
            Case "Libras"
                      Fecha = Me.TxtFecha.Value
                      Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      Me.DtaTasas.Refresh
                If Not DtaTasas.Recordset.EOF Then
                 MontoTasa = Me.DtaTasas.Recordset("MontoLibras")
               ' MontoCheque = MontoTasa * MontoCheque
                Else
               'MontoCheque = 1 * MontoCheque
                End If
         
         End Select

End If

TotalCredito = TotalCredito + MontoCheque
Me.TxtCredito.Text = Format(TotalCredito, "##,##0.00")
Me.TxtDebito.Text = Format(TotalDebito, "##,##0.00")
Me.TxtDiferencia.Text = Format(TotalDebito - TotalCredito, "##,##0.00")



 Me.DBCodigo.Enabled = False
 'If Cheque = True Then
    'TxtLetras.Text = ew.ConvertCurrencyToEnglish(TxtMonto.Text, "Cordobas")
   If TipoMoneda = "D�lares" Then
    TxtLetras.Text = sw.ConvertCurrencyToSpanish(TxtMonto.Text, "D�lares")
    
   
   ElseIf TipoMoneda = "C�rdobas" Then
    'TxtLetras.Text = ew.ConvertCurrencyToEnglish(TxtMonto.Text, "Cordobas")
    TxtLetras.Text = sw.ConvertCurrencyToSpanish(Me.TxtMonto.Text, "C�rdobas")
    
    
   End If

 'End If


End If

Exit Sub
TipoErrs:
ControlErrores

End Sub

Private Sub TxtCodigoEmpleado_Change()
On Error GoTo TipoErrs
Dim MontoTasa As Double, Fecha As Long
Dim SQL As String
Criterio = "CodCuentas='" & Me.TxtCodigoEmpleado.Text & "'"
If Me.DtaCuentas.Recordset.RecordCount > 0 Then Me.DtaCuentas.Recordset.MoveFirst
Me.DtaCuentas.Recordset.Find (Criterio)
  

If Not DtaCuentas.Recordset.EOF Then

'////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////CARGO LOS CHEQUES PENDIENTES/////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////
SQL = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, " & _
"Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, " & _
"Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, " & _
"Transacciones.NumeroMovimiento , Periodos.Periodo, Transacciones.Beneficiario " & _
"FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo " & _
"WHERE (Transacciones.ChequeNo = '#######')  AND (Transacciones.CodCuentas = '" & Me.DBCodigo.Text & "' ) AND " & _
"(Transacciones.DescripcionMovimiento <> '**********CANCELADO*************') ORDER BY Transacciones.NTransaccion"

Me.AdoPendientes.RecordSource = SQL
Me.AdoPendientes.Refresh





   Me.TxtNombreBanco.Text = Me.DtaCuentas.Recordset("DescripcionCuentas")
              TipoMoneda = Me.DtaCuentas.Recordset("TipoMoneda")
              
          Me.CmbMoneda.Text = TipoMoneda

         Select Case TipoMoneda
            Case "C�rdobas"
                      Fecha = Me.TxtFecha.Value
                      Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      Me.DtaTasas.Refresh
                If Not DtaTasas.Recordset.EOF Then
                 MontoTasa = Me.DtaTasas.Recordset("MontoCordobas")
                 If MontoTasa = 0 Then
                   MontoTasa = Format(1, "##,##0.000000")
                 End If
                 Me.LblTasa.Caption = Format((1 / MontoTasa), "##,##0.000000")
                Else
                 Me.LblTasa.Caption = Format(1, "##,##0.000000")
                End If
            
            Case "D�lares"
                    Me.LblTasa.Caption = Format(1, "##,##0.000000")
            
            Case "Libras"
                      Fecha = Me.TxtFecha.Value
                      Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      Me.DtaTasas.Refresh
                If Not DtaTasas.Recordset.EOF Then
                 MontoTasa = Me.DtaTasas.Recordset("MontoLibras")
                 Me.LblTasa.Caption = Format(MontoTasa, "##,##0.000000")
                Else
                Me.LblTasa.Caption = Format(1, "##,##0.000000")
                End If
         
         End Select



 Me.TxtFecha.Enabled = True
 If Not IsNull(Me.DtaBancos.Recordset("TipoCuenta")) Then
   TipoCuenta = Me.DtaBancos.Recordset("TipoCuenta")
 End If
 
 If TipoMoneda = "C�rdobas" Then
  Me.LblSimbolo.Caption = "Monto $"
 ElseIf TipoMoneda = "D�lares" Then
  Me.LblSimbolo.Caption = "Monto C$"
 End If
 
 
ElseIf Not Me.DBCodigo.Text = "" Then
  'MsgBox "El codigo Digitado no es correcto", vbCritical, "Sistema Contable"
  Exit Sub
Else
  Me.TxtNombreBanco.Text = ""
End If
Exit Sub
TipoErrs:
MsgBox err.Description

End Sub

Private Sub TxtFecha_Change()
On Error GoTo TipoErrs
Dim FEC As Date
Dim MontoTasa As Double, Fecha As Long
Criterio = "CodCuentas='" & Me.DBCodigo.Text & "'"
Me.DtaCuentas.Recordset.Find (Criterio)
If Not DtaCuentas.Recordset.EOF Then
              TipoMoneda = Me.DtaCuentas.Recordset("TipoMoneda")

         Select Case TipoMoneda
            Case "C�rdobas"
                      Fecha = Me.TxtFecha.Value
                      Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      Me.DtaTasas.Refresh
                If Not DtaTasas.Recordset.EOF Then
                 MontoTasa = Me.DtaTasas.Recordset("MontoCordobas")
                 If MontoTasa = 0 Then
                   MontoTasa = Format(1, "##,##0.000000")
                 End If
                 Me.LblTasa.Caption = Format((1 / MontoTasa), "##,##0.000000")
                Else
                 Me.LblTasa.Caption = Format(1, "##,##0.000000")
                End If
            
            Case "D�lares"
                    Me.LblTasa.Caption = Format(1, "##,##0.000000")
            
            Case "Libras"
                      Fecha = Me.TxtFecha.Value
                      Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      Me.DtaTasas.Refresh
                If Not DtaTasas.Recordset.EOF Then
                 MontoTasa = Me.DtaTasas.Recordset("MontoLibras")
                 Me.LblTasa.Caption = Format(MontoTasa, "##,##0.000000")
                Else
                Me.LblTasa.Caption = Format(1, "##,##0.000000")
                End If
         
         End Select
  End If
 Me.DBGTransacciones.Enabled = True
 mes = Month(Me.TxtFecha.Value)
 A�o = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(A�o, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  Me.DBGTransacciones.Enabled = True
  Me.TxtMonto.Enabled = True
  Me.TxtNombre.Enabled = True
  Me.TxtMemo.Enabled = True
  Me.TxtPeriodo.Text = DtaConsulta.Recordset("Periodo")
  NumeroPeriodo = DtaConsulta.Recordset("NPeriodo")
  NumeroTransaccion = DtaConsulta.Recordset("NTransacciones")
  EstadoPeriodo = DtaConsulta.Recordset("EstadoPeriodo")
  Me.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito From Transacciones GROUP BY Transacciones.CodCuentas Having (((Transacciones.CodCuentas) = '" & Me.DBCodigo.Text & "'))"
  Me.DtaHistorial.Refresh
  If Not Me.DtaHistorial.Recordset.EOF Then
   If Not IsNull(Me.DtaHistorial.Recordset("MDebito")) Then
    Debito = Me.DtaHistorial.Recordset("MDebito")
   Else
    Debito = 0
   End If
   If Not IsNull(Me.DtaHistorial.Recordset("MCredito")) Then
    Credito = Me.DtaHistorial.Recordset("MCredito")
   Else
    Credito = 0
   End If
   Me.TxtSaldoActual.Text = Format(Debito - Credito, "##,##0.00")
  End If
  
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
 If Not IsNull(DtaTasas.Recordset("FechaTasas")) Then
  FEC = Format(DtaTasas.Recordset("FechaTasas"), "dd/mm/yyyy")
  Fecha = FEC
 End If
   
    Encontrado = True
    Cambio = DtaTasas.Recordset("MontoCordobas")
    MDIPrimero.StatusBar2.Panels(2) = "Tasa Dolar: " & Format(Cambio, "##,##0.00") & "    " & "Tasa Libras: " & Format(DtaTasas.Recordset("MontoLibras"), "##,##0.00")
End If

If Not Encontrado Then
  MsgBox "La tasa de esta Fecha no ha sido Grabada"
  'Cancel = 100
  Tasa = False
  frmTasa2.Show 1
End If



  Me.DBGTransacciones.Columns(0).Button = True
    Me.DBGTransacciones.Columns(1).Locked = True
  Me.DBGTransacciones.Columns(6).Button = True
  Me.DBGTransacciones.Columns(6).Locked = True
  Me.DBGTransacciones.Columns(0).Width = 1500
  Me.DBGTransacciones.Columns(2).Width = 1000
  Me.DBGTransacciones.Columns(3).Caption = "Descripcion"
  Me.DBGTransacciones.Columns(4).Width = 1000
  Me.DBGTransacciones.Columns(4).Button = True
  Me.DBGTransacciones.Columns(5).Width = 1000
  Me.DBGTransacciones.Columns(6).Width = 800
  Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
  Me.DBGTransacciones.Columns(7).Locked = True
  Me.DBGTransacciones.Columns(7).Width = 1200
  Me.DBGTransacciones.Columns(7).NumberFormat = "##,##0.000000"
  Me.DBGTransacciones.Columns(8).Width = 1200
  Me.DBGTransacciones.Columns(8).NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns(9).Width = 1200
  Me.DBGTransacciones.Columns(9).NumberFormat = "##,##0.00"
  Me.DBGTransacciones.Columns(10).Visible = False
  Me.DBGTransacciones.Columns(11).Visible = False
  Me.DBGTransacciones.Columns(12).Visible = False
  Me.DBGTransacciones.Columns(11).Visible = False
  Me.DBGTransacciones.Columns(12).Visible = False
  Me.DBGTransacciones.Columns(13).Visible = False
  Me.DBGTransacciones.Columns(14).Visible = False
  Me.DBGTransacciones.Columns(15).Visible = False
    Me.DBGTransacciones.Columns(17).Visible = False
  Me.DBGTransacciones.Columns(18).Visible = False
  'Me.DBGTransacciones.Columns(19).Visible = False
  'Me.DBGTransacciones.Columns(16).Visible = False

Exit Sub
TipoErrs:
 ControlErrores

End Sub

Private Sub TxtFecha_GotFocus()
Dim Fechas1 As String
On Error GoTo TipoErrs
 Me.DBGTransacciones.Enabled = True
 mes = Month(Me.TxtFecha.Value)
 A�o = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(A�o, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
  Fechas1 = Format(FechaIni, "yyyy/mm/dd")
 Fechas2 = Format(FechaFin, "yyyy/mm/dd")
 Me.DtaConsulta.RecordSource = "SELECT  NPeriodo, NumeroTabla, FechaPeriodo, EstadoPeriodo, NTransacciones, Periodo From Periodos WHERE     (FechaPeriodo BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas2 & "', 102))"
' Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  Me.DBGTransacciones.Enabled = True
 Me.TxtMonto.Enabled = True
 Me.TxtNombre.Enabled = True
 Me.TxtMemo.Enabled = True
  Me.TxtPeriodo.Text = DtaConsulta.Recordset("Periodo")
  NumeroPeriodo = DtaConsulta.Recordset("NPeriodo")
  NumeroTransaccion = DtaConsulta.Recordset("NTransacciones")
  EstadoPeriodo = DtaConsulta.Recordset("EstadoPeriodo")
  Me.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito From Transacciones GROUP BY Transacciones.CodCuentas Having (((Transacciones.CodCuentas) = '" & Me.DBCodigo.Text & "'))"
  Me.DtaHistorial.Refresh
  If Not Me.DtaHistorial.Recordset.EOF Then
   If Not IsNull(Me.DtaHistorial.Recordset("MDebito")) Then
    Debito = Me.DtaHistorial.Recordset("MDebito")
   Else
    Debito = 0
   End If
   If Not IsNull(Me.DtaHistorial.Recordset("MCredito")) Then
    Credito = Me.DtaHistorial.Recordset("MCredito")
   Else
    Credito = 0
   End If
   Me.TxtSaldoActual.Text = Format(Debito - Credito, "##,##0.00")
  End If
  
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
NumFecha = Me.TxtFecha.Value
Fechas1 = Format(Me.TxtFecha.Value, "yyyy/mm/dd")
Me.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas1 & "', 102)) ORDER BY FechaTasas"
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
  'Cancel = 100
  Tasa = False
  frmTasa2.Show 1
End If

Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub TxtFecha_LostFocus()
On Error GoTo TipoErrs
Dim NumFecha As Long, Fechas1 As String, Fechas2 As String
mes = Month(Me.TxtFecha.Value)
 A�o = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(A�o, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Fechas1 = Format(FechaIni, "yyyy/mm/dd")
 Fechas2 = Format(FechaFin, "yyyy/mm/dd")
  Me.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito From Transacciones GROUP BY Transacciones.CodCuentas Having (((Transacciones.CodCuentas) = '" & Me.DBCodigo.Text & "'))"
  Me.DtaHistorial.Refresh
  If Not Me.DtaHistorial.Recordset.EOF Then
   If Not IsNull(Me.DtaHistorial.Recordset("MDebito")) Then
    Debito = Me.DtaHistorial.Recordset("MDebito")
   Else
    Debito = 0
   End If
   If Not IsNull(Me.DtaHistorial.Recordset("MCredito")) Then
    Credito = Me.DtaHistorial.Recordset("MCredito")
   Else
    Credito = 0
   End If
   Me.TxtSaldoActual.Text = Format(Debito - Credito, "##,##0.00")
  End If
 
 Me.DtaConsulta.RecordSource = "SELECT  NPeriodo, NumeroTabla, FechaPeriodo, EstadoPeriodo, NTransacciones, Periodo From Periodos WHERE     (FechaPeriodo BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas2 & "', 102))"
' Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  Me.TxtPeriodo.Text = DtaConsulta.Recordset("Periodo")
  Me.DBGTransacciones.Enabled = True
 Me.TxtMonto.Enabled = True
 Me.TxtNombre.Enabled = True
 Me.TxtMemo.Enabled = True
 End If
 
 '///////Verifico si esta registrada la fecha de la tasa//////

NumFecha = Me.TxtFecha.Value
Fechas1 = Format(Me.TxtFecha.Value, "yyyy/mm/dd")
Me.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas1 & "', 102)) ORDER BY FechaTasas"
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
  frmTasa2.Show 1
End If
 
 Exit Sub
TipoErrs:
 MsgBox err.Description
End Sub

Private Sub TxtMonto_Change()
'On Error GoTo TipoErrs
Dim SqlContratista As String, Fecha As Long

 Debito = 0
 Credito = 0
 TotalDebito = 0
 TotalCredito = 0

If Not DBCodigo.Text = "" Then
  If TxtMonto.Text = "" Then
   Credito = 0
  Else
   Credito = Me.TxtMonto
  End If

      NumFecha1 = Me.TxtFecha.Value
      NMovimiento = Val(Me.TxtNTransacciones)
'      Me.DtaConsulta.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.CodCuentas, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, TCambio*Debito AS MDebito, TCambio*Credito AS MCredito, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito From Transacciones WHERE (((Transacciones.FechaTransaccion)=" & NumFecha1 & ") AND ((Transacciones.NumeroMovimiento)=" & NMovimiento & "))"
      Me.DtaConsulta.RecordSource = "SELECT FechaTransaccion, CodCuentas, NTransaccion, NumeroMovimiento, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, TCambio, Debito, Credito From Transacciones  " & _
                                    "WHERE (FechaTransaccion = CONVERT(DATETIME, '" & Format(Me.TxtFecha.Value, "yyyy-mm-dd") & "', 102)) AND (NumeroMovimiento = " & NMovimiento & ") AND (CodCuentas <> '" & Me.DBCodigo.Text & "')"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
       If Me.TxtMonto.Text = "" Then
        MontoCheque = 0
       Else
         MontoCheque = Me.TxtMonto
       End If
       Debito = Me.DtaConsulta.Recordset("Debito")
       Credito = Me.DtaConsulta.Recordset("Credito")
       TotalDebito = TotalDebito + Debito
       TotalCredito = TotalCredito + Credito
       DtaConsulta.Recordset.MoveNext
      Loop
      

      
       Criterio = "CodCuentas='" & Me.DBCodigo.Text & "'"
       Me.DtaCuentas.Recordset.Find (Criterio)
       If Not DtaCuentas.Recordset.EOF Then
'                TipoMoneda = DtaCuentas.Recordset("TipoMoneda")
                TipoMoneda = Me.CmbMoneda.Text
                TipoCuenta = DtaCuentas.Recordset("TipoCuenta")
      If Me.TxtMonto.Text = "" Then
         MontoCheque = 0
      Else
           MontoCheque = Me.TxtMonto.Text
      End If
     
         Select Case TipoMoneda
            Case "C�rdobas"
                      Fecha = Me.TxtFecha.Value
                      Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      Me.DtaTasas.Refresh
                If Not DtaTasas.Recordset.EOF Then
                 MontoTasa = Me.DtaTasas.Recordset("MontoCordobas")
                 If MontoTasa = 0 Then
                   MontoTasa = 1
                 End If
                 'MontoCheque = (1 / MontoTasa) * MontoCheque
                Else
                 'MontoCheque = 1 * MontoCheque
                End If
            
            Case "D�lares"
                '  MontoCheque = 1 * MontoCheque
            
            Case "Libras"
                      Fecha = Me.TxtFecha.Value
                      Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      Me.DtaTasas.Refresh
                If Not DtaTasas.Recordset.EOF Then
                 MontoTasa = Me.DtaTasas.Recordset("MontoLibras")
               ' MontoCheque = MontoTasa * MontoCheque
                Else
               'MontoCheque = 1 * MontoCheque
                End If
         
         End Select

End If

TotalCredito = TotalCredito + MontoCheque
Me.TxtCredito.Text = Format(TotalCredito, "##,##0.00")
Me.TxtDebito.Text = Format(TotalDebito, "##,##0.00")
Me.TxtDiferencia.Text = Format(TotalDebito - TotalCredito, "##,##0.00")



 Me.DBCodigo.Enabled = False
 'If Cheque = True Then
    'TxtLetras.Text = ew.ConvertCurrencyToEnglish(TxtMonto.Text, "Cordobas")
   If TipoMoneda = "D�lares" Then
    TxtLetras.Text = sw.ConvertCurrencyToSpanish(TxtMonto.Text, "D�lares")
    
   
   ElseIf TipoMoneda = "C�rdobas" Then
    'TxtLetras.Text = ew.ConvertCurrencyToEnglish(TxtMonto.Text, "Cordobas")
    TxtLetras.Text = sw.ConvertCurrencyToSpanish(TxtMonto.Text, "C�rdobas")
    
   End If

 'End If


End If

Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub TxtSubTotal_Change()
  Me.TxtMonto.Text = CalcularMonto
End Sub


 Function CalcularMonto() As Double
  Dim SubTotal As Double, Iva As Double, Retenciones As Double, P1 As Double, P2 As Double, P3 As Double, P7 As Double, P10 As Double, P15 As Double
  
  If Me.Chk15.Value = 1 Then
   P15 = 0.15
  Else
   P15 = 0
  End If
  
  If Me.TxtSubTotal.Text <> "" Then
   If IsNumeric(Me.TxtSubTotal.Text) Then
    SubTotal = CDbl(Me.TxtSubTotal.Text)
   Else
    Me.TxtSubTotal.Text = ""
    SubTotal = 0
   End If
  Else
    SubTotal = 0
  End If
  
'   Me.TxtIVa.Text = SubTotal * P15
  
  If Me.Chk1.Value = 1 Then
   P1 = 0.01
  Else
   P1 = 0
  End If
  
  If Me.Chk2.Value = 1 Then
   P2 = 0.02
  Else
   P2 = 0
  End If
  
  If Me.Chk3.Value = 1 Then
   P3 = 0.03
  Else
   P3 = 0
  End If
  
  If Me.Chk7.Value = 1 Then
   P7 = 0.07
  Else
   P7 = 0
  End If
  
  If Me.Chk10.Value = 1 Then
   P10 = 0.1
  Else
   P10 = 0
  End If
  
  Iva = SubTotal * P15
  Me.TxtIVa.Text = Format(Iva, "##,##0.00")

  
  Retenciones = SubTotal * P1 + SubTotal * P2 + SubTotal * P3 + SubTotal * P7 + SubTotal * P10
  Me.TxtRetenciones.Text = Format(Retenciones, "##,##0.00")
  
  CalcularMonto = Format(SubTotal + Iva - Retenciones, "##,##0.00")
 
 End Function


