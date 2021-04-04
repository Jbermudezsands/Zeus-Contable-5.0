VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FrmReparar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auditor - Reparar"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   11790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10440
      TabIndex        =   7
      Top             =   6720
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc AdoIndices 
      Height          =   375
      Left            =   600
      Top             =   7560
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
      Caption         =   "AdoIndices"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Reparar Indices"
      TabPicture(0)   =   "FrmReparar.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DBGCuentas"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "Picture1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Reparar Transacciones"
      TabPicture(1)   =   "FrmReparar.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "SkinLabel15"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "SkinLabel14"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "SkinLabel13"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "DBGTransacciones"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "TxtDiferencia"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "TxtDebito"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "TxtCredito"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
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
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "0.00"
         Top             =   5640
         Width           =   1335
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
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "0.00"
         Top             =   5640
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "0.00"
         Top             =   5640
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   -74760
         ScaleHeight     =   4305
         ScaleWidth      =   11025
         TabIndex        =   21
         Top             =   1920
         Visible         =   0   'False
         Width           =   11055
         Begin VB.CommandButton CmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   5640
            TabIndex        =   36
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CommandButton CmdGrabar 
            Caption         =   "Grabar"
            Height          =   375
            Left            =   3360
            TabIndex        =   35
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Frame Frame3 
            Height          =   3135
            Left            =   2280
            TabIndex        =   22
            Top             =   480
            Width           =   5895
            Begin VB.TextBox TxtFuente 
               Height          =   285
               Left            =   1440
               TabIndex        =   34
               Top             =   2640
               Width           =   1815
            End
            Begin VB.ComboBox CmbMoneda 
               Height          =   315
               ItemData        =   "FrmReparar.frx":0038
               Left            =   1440
               List            =   "FrmReparar.frx":0042
               TabIndex        =   31
               Text            =   "Córdobas"
               Top             =   2280
               Width           =   1815
            End
            Begin VB.TextBox TxtDescripcionIndice 
               Height          =   495
               Left            =   1440
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   30
               Top             =   1560
               Width           =   3975
            End
            Begin VB.TextBox TxtMovimiento2 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1440
               TabIndex        =   24
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox TxtPeriodo2 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1440
               TabIndex        =   23
               Top             =   960
               Width           =   1095
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
               Height          =   255
               Left            =   840
               OleObjectBlob   =   "FrmReparar.frx":0059
               TabIndex        =   25
               Top             =   240
               Width           =   615
            End
            Begin MSComCtl2.DTPicker TxtFechaIndice2 
               Height          =   285
               Left            =   1440
               TabIndex        =   26
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   503
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   17104897
               CurrentDate     =   39117
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmReparar.frx":00C1
               TabIndex        =   27
               Top             =   600
               Width           =   1215
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
               Height          =   255
               Left            =   720
               OleObjectBlob   =   "FrmReparar.frx":013B
               TabIndex        =   28
               Top             =   960
               Width           =   615
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmReparar.frx":01A9
               TabIndex        =   29
               Top             =   1320
               Width           =   2055
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
               Height          =   255
               Left            =   360
               OleObjectBlob   =   "FrmReparar.frx":0235
               TabIndex        =   32
               Top             =   2280
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
               Height          =   255
               Left            =   840
               OleObjectBlob   =   "FrmReparar.frx":02A9
               TabIndex        =   33
               Top             =   2640
               Width           =   495
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Consulta de Indices"
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   11295
         Begin VB.TextBox TxtPeriodoTransacciones 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6120
            TabIndex        =   19
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox TxtMovimientoTransacciones 
            Height          =   285
            Left            =   3720
            TabIndex        =   15
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton CmdConsultarTransaccion 
            Caption         =   "Consultar"
            Height          =   375
            Left            =   8520
            TabIndex        =   9
            Top             =   360
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmReparar.frx":0313
            TabIndex        =   10
            Top             =   360
            Width           =   495
         End
         Begin MSComCtl2.DTPicker DTPFechaFin 
            Height          =   285
            Left            =   720
            TabIndex        =   11
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Format          =   17104897
            CurrentDate     =   39117
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   2520
            OleObjectBlob   =   "FrmReparar.frx":037D
            TabIndex        =   16
            Top             =   360
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   5520
            OleObjectBlob   =   "FrmReparar.frx":03F7
            TabIndex        =   20
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Consulta de Indices"
         Height          =   975
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   11295
         Begin VB.TextBox TxtPeriodo 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6720
            TabIndex        =   18
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox TxtMovimientoIndice 
            Height          =   285
            Left            =   4080
            TabIndex        =   13
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton CmdConsultaIndices 
            Caption         =   "Consultar"
            Height          =   375
            Left            =   8640
            TabIndex        =   5
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton CmdAgregar 
            Caption         =   "Agregar"
            Height          =   375
            Left            =   9960
            TabIndex        =   4
            Top             =   360
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   360
            OleObjectBlob   =   "FrmReparar.frx":0465
            TabIndex        =   2
            Top             =   360
            Width           =   615
         End
         Begin MSComCtl2.DTPicker DTPFechaIni 
            Height          =   285
            Left            =   960
            TabIndex        =   3
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            Format          =   17104897
            CurrentDate     =   39117
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   2880
            OleObjectBlob   =   "FrmReparar.frx":04CD
            TabIndex        =   14
            Top             =   360
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   6120
            OleObjectBlob   =   "FrmReparar.frx":0547
            TabIndex        =   17
            Top             =   360
            Width           =   735
         End
      End
      Begin TrueOleDBGrid80.TDBGrid DBGCuentas 
         Bindings        =   "FrmReparar.frx":05B5
         Height          =   4695
         Left            =   -74880
         TabIndex        =   6
         Top             =   1560
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   8281
         _LayoutType     =   1
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "FechaTransaccion"
         Columns(0).DataField=   "FechaTransaccion"
         Columns(0).DataWidth=   19
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "NumeroMovimiento"
         Columns(1).DataField=   "NumeroMovimiento"
         Columns(1).DataWidth=   11
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "DescripcionMovimiento"
         Columns(2).DataField=   "DescripcionMovimiento"
         Columns(2).DataWidth=   255
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Nperiodo"
         Columns(3).DataField=   "Nperiodo"
         Columns(3).DataWidth=   11
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Fuente"
         Columns(4).DataField=   "Fuente"
         Columns(4).DataWidth=   255
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "TipoMoneda"
         Columns(5).DataField=   "TipoMoneda"
         Columns(5).DataWidth=   50
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).Caption=   "Movimientos de Indices"
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3096"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3016"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=131588"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(0)._AlignLeft=0"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3069"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2990"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=131588"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(1)._AlignLeft=0"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=3704"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3625"
         Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=131588"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=1826"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1746"
         Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=131588"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(3)._AlignLeft=0"
         Splits(0)._ColumnProps(24)=   "Column(4).Width=3254"
         Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=3175"
         Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=131588"
         Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(29)=   "Column(5).Width=3254"
         Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=3175"
         Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=131588"
         Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
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
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HFFAEFF&,.fgcolor=&H800080&"
         _StyleDefs(20)  =   ":id=22,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(21)  =   ":id=22,.fontname=Lucida Calligraphy"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bgcolor=&HECB877&"
         _StyleDefs(23)  =   ":id=14,.fgcolor=&H800000&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(24)  =   ":id=14,.strikethrough=0,.charset=0"
         _StyleDefs(25)  =   ":id=14,.fontname=MS Sans Serif"
         _StyleDefs(26)  =   "Splits(0).FooterStyle:id=15,.parent=3,.alignment=2,.bgcolor=&HFF0000&"
         _StyleDefs(27)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(28)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(29)  =   "Splits(0).EditorStyle:id=17,.parent=7"
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
      Begin TrueOleDBGrid80.TDBGrid DBGTransacciones 
         Bindings        =   "FrmReparar.frx":05CE
         Height          =   3975
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7011
         _LayoutType     =   1
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "CodCuentas"
         Columns(0).DataField=   "CodCuentas"
         Columns(0).DataWidth=   50
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "NombreCuenta"
         Columns(1).DataField=   "NombreCuenta"
         Columns(1).DataWidth=   255
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "FechaTransaccion"
         Columns(2).DataField=   "FechaTransaccion"
         Columns(2).DataWidth=   19
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "NPeriodo"
         Columns(3).DataField=   "NPeriodo"
         Columns(3).DataWidth=   11
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "NumeroMovimiento"
         Columns(4).DataField=   "NumeroMovimiento"
         Columns(4).DataWidth=   11
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "VoucherNo"
         Columns(5).DataField=   "VoucherNo"
         Columns(5).DataWidth=   50
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "DescripcionMovimiento"
         Columns(6).DataField=   "DescripcionMovimiento"
         Columns(6).DataWidth=   255
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Clave"
         Columns(7).DataField=   "Clave"
         Columns(7).DataWidth=   10
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "TCambio"
         Columns(8).DataField=   "TCambio"
         Columns(8).DataWidth=   23
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Debito"
         Columns(9).DataField=   "Debito"
         Columns(9).DataWidth=   22
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Credito"
         Columns(10).DataField=   "Credito"
         Columns(10).DataWidth=   22
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   11
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).Caption=   "Movimientos de Indices"
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=11"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3254"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=131588"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3254"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=131588"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=3096"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3016"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=131588"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(2)._AlignLeft=0"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=1826"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1746"
         Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=131588"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(3)._AlignLeft=0"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=3069"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2990"
         Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=131588"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(28)=   "Column(4)._AlignLeft=0"
         Splits(0)._ColumnProps(29)=   "Column(5).Width=3254"
         Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=3175"
         Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=131588"
         Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(34)=   "Column(6).Width=3704"
         Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=3625"
         Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=131588"
         Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(39)=   "Column(7).Width=1667"
         Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=1588"
         Splits(0)._ColumnProps(42)=   "Column(7)._ColStyle=131588"
         Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(44)=   "Column(8).Width=3254"
         Splits(0)._ColumnProps(45)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(46)=   "Column(8)._WidthInPix=3175"
         Splits(0)._ColumnProps(47)=   "Column(8)._ColStyle=131588"
         Splits(0)._ColumnProps(48)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(49)=   "Column(8)._AlignLeft=0"
         Splits(0)._ColumnProps(50)=   "Column(9).Width=3254"
         Splits(0)._ColumnProps(51)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(52)=   "Column(9)._WidthInPix=3175"
         Splits(0)._ColumnProps(53)=   "Column(9)._ColStyle=131588"
         Splits(0)._ColumnProps(54)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(55)=   "Column(9)._AlignLeft=0"
         Splits(0)._ColumnProps(56)=   "Column(10).Width=3254"
         Splits(0)._ColumnProps(57)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(58)=   "Column(10)._WidthInPix=3175"
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
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HFFAEFF&,.fgcolor=&H800080&"
         _StyleDefs(20)  =   ":id=22,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(21)  =   ":id=22,.fontname=Lucida Calligraphy"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bgcolor=&HECB877&"
         _StyleDefs(23)  =   ":id=14,.fgcolor=&H800000&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(24)  =   ":id=14,.strikethrough=0,.charset=0"
         _StyleDefs(25)  =   ":id=14,.fontname=MS Sans Serif"
         _StyleDefs(26)  =   "Splits(0).FooterStyle:id=15,.parent=3,.alignment=2,.bgcolor=&HFF0000&"
         _StyleDefs(27)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(28)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(29)  =   "Splits(0).EditorStyle:id=17,.parent=7"
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   10080
         OleObjectBlob   =   "FrmReparar.frx":05ED
         TabIndex        =   37
         Top             =   6120
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   8640
         OleObjectBlob   =   "FrmReparar.frx":0659
         TabIndex        =   38
         Top             =   6120
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmReparar.frx":06C3
         TabIndex        =   42
         Top             =   5640
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc AdoTransacciones 
      Height          =   375
      Left            =   600
      Top             =   8040
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
   Begin MSAdodcLib.Adodc AdoConsulta 
      Height          =   375
      Left            =   3840
      Top             =   8040
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
End
Attribute VB_Name = "FrmReparar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdConsultar_Click()
  
End Sub

Private Sub CmdConsultarIndices_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub CmdAgregar_Click()
Me.Picture1.Visible = True
Me.TxtFechaIndice2.Value = Me.DTPFechaIni.Value
Me.TxtPeriodo2.Text = Me.TxtPeriodo.Text
Me.TxtMovimiento2.Text = Me.TxtMovimientoIndice.Text

End Sub

Private Sub CmdCancelar_Click()
Me.Picture1.Visible = False
End Sub

Private Sub CmdConsultaIndices_Click()
Dim Fechas1 As String, Fechas2 As String

'////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////BUSCO EL PERIODO DE LA TRANSACCION ////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////
 mes = Month(Me.DTPFechaIni.Value)
 Año = Year(Me.DTPFechaIni.Value)
 FechaIni = CDate("1/" & Month(Me.DTPFechaIni.Value) & "/" & Year(Me.DTPFechaIni.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 Fechas1 = Format(FechaIni, "yyyy/mm/dd")
 Fechas2 = Format(FechaFin, "yyyy/mm/dd")
 
  Me.AdoConsulta.RecordSource = "SELECT NPeriodo, NumeroTabla, FechaPeriodo, EstadoPeriodo, NTransacciones, Periodo From Periodos WHERE     (FechaPeriodo BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas2 & "', 102))"
  Me.AdoConsulta.Refresh
 If Not AdoConsulta.Recordset.EOF Then
  Me.TxtPeriodo.Text = AdoConsulta.Recordset("NPeriodo")
 End If
 
 Me.AdoIndices.RecordSource = "SELECT * From IndiceTransaccion Where (NumeroMovimiento = " & Me.TxtMovimientoIndice.Text & ") And (NPeriodo = " & Me.TxtPeriodo.Text & ")"
 Me.AdoIndices.Refresh
End Sub

Private Sub CmdConsultarTransaccion_Click()
Dim R As Variant
'////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////BUSCO EL PERIODO DE LA TRANSACCION ////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////
 mes = Month(Me.DTPFechaFin.Value)
 Año = Year(Me.DTPFechaFin.Value)
 FechaIni = CDate("1/" & Month(Me.DTPFechaFin.Value) & "/" & Year(Me.DTPFechaFin.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 Fechas1 = Format(FechaIni, "yyyy/mm/dd")
 Fechas2 = Format(FechaFin, "yyyy/mm/dd")
 
  Me.AdoConsulta.RecordSource = "SELECT NPeriodo, NumeroTabla, FechaPeriodo, EstadoPeriodo, NTransacciones, Periodo From Periodos WHERE  (FechaPeriodo BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas2 & "', 102))"
  Me.AdoConsulta.Refresh
 If Not AdoConsulta.Recordset.EOF Then
  Me.TxtPeriodoTransacciones.Text = AdoConsulta.Recordset("NPeriodo")
 End If
 
 Me.AdoTransacciones.RecordSource = "SELECT CodCuentas, NombreCuenta, FechaTransaccion, NPeriodo, NumeroMovimiento, VoucherNo, DescripcionMovimiento, Clave, TCambio, Debito, Credito From Transacciones Where (NPeriodo = " & Me.TxtPeriodoTransacciones.Text & ") And (NumeroMovimiento = " & Me.TxtMovimientoTransacciones.Text & ")"
 Me.AdoTransacciones.Refresh
 
 R = SumasDebitos(Me.TxtMovimientoTransacciones.Text, Me.TxtPeriodoTransacciones.Text)

 
End Sub

Private Sub Command1_Click()

End Sub

Private Sub CmdGrabar_Click()
Dim Fechas1 As String, Fechas2 As String

On Error GoTo TipoErrs
'////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////BUSCO EL PERIODO DE LA TRANSACCION ////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////
 mes = Month(Me.DTPFechaIni.Value)
 Año = Year(Me.DTPFechaIni.Value)
 FechaIni = CDate("1/" & Month(Me.DTPFechaIni.Value) & "/" & Year(Me.DTPFechaIni.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 Fechas1 = Format(FechaIni, "yyyy/mm/dd")
 Fechas2 = Format(FechaFin, "yyyy/mm/dd")
 
  Me.AdoConsulta.RecordSource = "SELECT NPeriodo, NumeroTabla, FechaPeriodo, EstadoPeriodo, NTransacciones, Periodo From Periodos WHERE     (FechaPeriodo BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas2 & "', 102))"
  Me.AdoConsulta.Refresh
 If Not AdoConsulta.Recordset.EOF Then
  Me.TxtPeriodo.Text = AdoConsulta.Recordset("NPeriodo")
 End If
 
 Me.AdoIndices.RecordSource = "SELECT * From IndiceTransaccion Where (NumeroMovimiento = " & Me.TxtMovimientoIndice.Text & ") And (NPeriodo = " & Me.TxtPeriodo.Text & ")"
 Me.AdoIndices.Refresh
 If Me.AdoIndices.Recordset.EOF Then
    Me.AdoIndices.Recordset.AddNew
      Me.AdoIndices.Recordset("FechaTransaccion") = Format(Me.DTPFechaIni.Value, "dd/mm/yyyy")
      Me.AdoIndices.Recordset("NumeroMovimiento") = Me.TxtMovimientoIndice.Text
      Me.AdoIndices.Recordset("DescripcionMovimiento") = Me.TxtDescripcionIndice.Text
      Me.AdoIndices.Recordset("Nperiodo") = Me.TxtPeriodo2.Text
      Me.AdoIndices.Recordset("Fuente") = Me.TxtFuente.Text
      Me.AdoIndices.Recordset("TipoMoneda") = Me.CmbMoneda.Text
    Me.AdoIndices.Recordset.Update
 Else
   MsgBox "Ya Existe esta transaccion en la tabla de Indices", vbCritical, "Sistema Contable"
   Me.Picture1.Visible = False
 End If
 
 Me.Picture1.Visible = False
 
 Exit Sub
TipoErrs:
 MsgBox err.Description
 
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()

End Sub

Private Sub DBGTransacciones_AfterUpdate()
Dim R As Variant

 R = SumasDebitos(Me.TxtMovimientoTransacciones.Text, Me.TxtPeriodoTransacciones.Text)
End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd
Me.SSTab1.BackColor = RGB(219, 226, 242)


Me.DTPFechaFin.Value = Format(Now, "dd/mm/yyyy")
Me.DTPFechaIni.Value = Format(Now, "dd/mm/yyyy")
 Me.DBGCuentas.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DBGCuentas.OddRowStyle.BackColor = &H80000005
 Me.DBGCuentas.AlternatingRowStyle = True
 
 Me.DBGTransacciones.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DBGTransacciones.OddRowStyle.BackColor = &H80000005
 Me.DBGTransacciones.AlternatingRowStyle = True
 
With Me.AdoIndices
   .ConnectionString = Conexion
End With

With Me.AdoTransacciones
   .ConnectionString = Conexion
End With

With Me.AdoConsulta
   .ConnectionString = Conexion
End With

End Sub
