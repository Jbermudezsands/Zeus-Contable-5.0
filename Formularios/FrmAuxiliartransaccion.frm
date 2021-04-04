VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form FrmTransacciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Transacciones"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   12945
   Begin VB.Data DtaNacceso 
      Caption         =   "DtaNacceso"
      Connect         =   ";DATABASENAME="" + Ruta + "";UID=Administrador;PWD=DFID"
      DatabaseName    =   "D:\DFID\dfid.bak"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Accesos"
      Top             =   7200
      Visible         =   0   'False
      Width           =   10515
   End
   Begin VB.Data DtaPeriodos 
      Caption         =   "DtaPeriodos"
      Connect         =   ";DATABASENAME="" + Ruta + "";UID=Administrador;PWD=DFID"
      DatabaseName    =   "D:\DFID\dfid.bak"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Periodos"
      Top             =   8160
      Visible         =   0   'False
      Width           =   4995
   End
   Begin VB.Data DtaHistorial 
      Caption         =   "DtaHistorial"
      Connect         =   ";DATABASENAME="" + Ruta + "";UID=Administrador;PWD=DFID"
      DatabaseName    =   "D:\DFID\dfid.bak"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8040
      Width           =   3375
   End
   Begin VB.Data DtaIndice 
      Caption         =   "DtaIndice"
      Connect         =   ";DATABASENAME="" + Ruta + "";UID=Administrador;PWD=DFID"
      DatabaseName    =   "D:\DFID\dfid.bak"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "IndiceTransaccion"
      Top             =   7920
      Width           =   3375
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   810
      ItemData        =   "FrmTransacciones.frx":0000
      Left            =   960
      List            =   "FrmTransacciones.frx":000A
      TabIndex        =   24
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data DtaCuentas 
      Caption         =   "DtaCuentas"
      Connect         =   ";DATABASENAME="" + Ruta + "";UID=Administrador;PWD=DFID"
      DatabaseName    =   "D:\DFID\dfid.bak"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Cuentas"
      Top             =   7080
      Visible         =   0   'False
      Width           =   4995
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
      Picture         =   "FrmTransacciones.frx":001F
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   360
      Width           =   375
   End
   Begin VB.Data DtaTransaccionesNuevas 
      Caption         =   "DtaTransaccionesNuevas"
      Connect         =   ";DATABASENAME="" + Ruta + "";UID=Administrador;PWD=DFID"
      DatabaseName    =   "D:\DFID\dfid.bak"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Transacciones"
      Top             =   6960
      Width           =   3375
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
      TabIndex        =   20
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
      TabIndex        =   17
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
      TabIndex        =   16
      Text            =   "0.00"
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Data DtaTasas 
      Caption         =   "DtaTasas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6960
      Visible         =   0   'False
      Width           =   4995
   End
   Begin VB.Data DtaConsulta 
      Caption         =   "DtaConsulta"
      Connect         =   ";DATABASENAME="" + Ruta + "";UID=Administrador;PWD=DFID"
      DatabaseName    =   "D:\DFID\dfid.bak"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7560
      Width           =   3375
   End
   Begin VB.Data DtaTransacciones 
      Caption         =   "DtaTransacciones"
      Connect         =   ";DATABASENAME="" + Ruta + "";UID=Administrador;PWD=DFID"
      DatabaseName    =   "D:\DFID\dfid.bak"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7080
      Width           =   3375
   End
   Begin TrueDBGrid70.TDBGrid DBGTransacciones 
      Bindings        =   "FrmTransacciones.frx":016D
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   12735
      _ExtentX        =   22463
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
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   10862530
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
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowAddNew     =   -1  'True
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   10862530
      RowDividerColor =   10862530
      RowSubDividerColor=   10862530
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12735
      Begin VB.ComboBox CmbMoneda 
         Height          =   315
         ItemData        =   "FrmTransacciones.frx":018C
         Left            =   8520
         List            =   "FrmTransacciones.frx":0199
         TabIndex        =   26
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TxtNTransacciones 
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtFuente 
         Height          =   285
         Left            =   11640
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtPeriodo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin MSComCtl2.DTPicker TxtFecha 
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Format          =   53673985
         CurrentDate     =   38008
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo Moneda"
         Height          =   255
         Left            =   7320
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Fuente"
         Height          =   255
         Left            =   10920
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Transaccion No."
         Height          =   255
         Left            =   4560
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Periodo"
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
   Begin SmartButtonProject.SmartButton CmdGrabar 
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      ForeColor       =   12582912
      Caption         =   "Grabar"
      Picture         =   "FrmTransacciones.frx":01B8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SmartButtonProject.SmartButton CmdAnterior 
      Height          =   855
      Left            =   3120
      TabIndex        =   10
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      ForeColor       =   12582912
      Caption         =   "Anterior"
      Picture         =   "FrmTransacciones.frx":0A92
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SmartButtonProject.SmartButton CmdSiguiente 
      Height          =   855
      Left            =   4560
      TabIndex        =   11
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      ForeColor       =   12582912
      Caption         =   "Siguiente"
      Picture         =   "FrmTransacciones.frx":0EE4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SmartButtonProject.SmartButton CmdBorrar 
      Height          =   855
      Left            =   6000
      TabIndex        =   12
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      ForeColor       =   12582912
      Caption         =   "Borrar"
      Picture         =   "FrmTransacciones.frx":1336
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SmartButtonProject.SmartButton CmdNuevo 
      Height          =   855
      Left            =   1680
      TabIndex        =   13
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      ForeColor       =   12582912
      Caption         =   "Nuevo"
      Picture         =   "FrmTransacciones.frx":1650
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SmartButtonProject.SmartButton CmdSalir 
      Height          =   855
      Left            =   11880
      TabIndex        =   14
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      ForeColor       =   12582912
      Caption         =   "Salir"
      Picture         =   "FrmTransacciones.frx":1AA2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SmartButtonProject.SmartButton SmartButton1 
      Height          =   855
      Left            =   7440
      TabIndex        =   15
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      ForeColor       =   12582912
      Caption         =   "Borrar Linea"
      Picture         =   "FrmTransacciones.frx":7D3C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label7 
      Caption         =   "Diferencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3120
      TabIndex        =   21
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Debito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   10080
      TabIndex        =   19
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Credito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   11400
      TabIndex        =   18
      Top             =   4560
      Width           =   1455
   End
End
Attribute VB_Name = "FrmTransacciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
 Mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente, IndiceTransaccion.TipoMoneda From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & " ) AND ((IndiceTransaccion.NumeroMovimiento)= " & NumeroTransaccion & "))"
 Me.DtaConsulta.Refresh
 If Not DtaTransacciones.Recordset.EOF Then
 Me.DtaTransacciones.Recordset.MoveFirst
 End If
       
       If Not DtaConsulta.Recordset.EOF Then
         If Not IsNull(Me.CmbMoneda.Text = Me.DtaConsulta.Recordset.TipoMoneda) Then
          'Me.CmbMoneda.Text = Me.DtaConsulta.Recordset.TipoMoneda
         End If
        If Not Me.DBGTransacciones.Columns(3).Text = "" Then
 
          Me.DtaConsulta.Recordset.Edit
          Me.DtaConsulta.Recordset.TipoMoneda = Me.CmbMoneda.Text
          Me.DtaConsulta.Recordset.DescripcionMovimiento = Me.DBGTransacciones.Columns(3).Text
          Me.DtaConsulta.Recordset.Update
        Else
          Me.DtaConsulta.Recordset.Edit
          'Me.DtaConsulta.Recordset.TipoMoneda = Me.CmbMoneda.Text
          Me.DtaConsulta.Recordset.Update
        End If
       End If
 TotalDiferencia = 0
 TotalCredito = 0
 TotalDebito = 0
 Debito = 0
 Credito = 0
 Diferencia = 0
 
If Me.TxtNTransacciones = 0 Then
 Mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones WHERE (((Transacciones.NombreCuenta)<>'**********CANCELADO*************') AND ((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ")) ORDER BY Transacciones.NumeroMovimiento"
 'Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & "))ORDER BY Transacciones.NumeroMovimiento"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
   NumeroTransaccion = DtaConsulta.Recordset.NumeroMovimiento
   Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.NombreCuenta)<>'**********CANCELADO*************') AND ((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion "
   'Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion"
   Me.DtaTransacciones.Refresh
   If Not DtaTransacciones.Recordset.EOF Then
     Me.TxtFecha.Value = Me.DtaTransacciones.Recordset.FechaTransaccion
     Me.TxtPeriodo.Text = Me.DtaTransacciones.Recordset.Periodo
     Me.TxtNTransacciones.Text = Me.DtaTransacciones.Recordset.NumeroMovimiento
     NumeroTransaccion = Me.DtaTransacciones.Recordset.NumeroMovimiento
     Me.TxtFuente.Text = Me.DtaTransacciones.Recordset.Fuente
     '//////Sumo los Totales/////////////////////
   
    Debito = 0
    Credito = 0
    TotalDebito = 0
    TotalCredito = 0
      NumFecha1 = Me.TxtFecha.Value
      NMovimiento = Val(Me.TxtNTransacciones)
      Me.DtaConsulta.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.CodCuentas, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, [Transacciones]![TCambio]*[Transacciones]![Debito] AS MDebito, [Transacciones]![TCambio]*[Transacciones]![Credito] AS MCredito, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito From Transacciones WHERE (((Transacciones.FechaTransaccion)=" & NumFecha1 & ") AND ((Transacciones.NumeroMovimiento)=" & NMovimiento & "))"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
       If Not IsNull(Me.DtaConsulta.Recordset.Debito) Then
       Debito = Me.DtaConsulta.Recordset.Debito
       End If
       If Not IsNull(Credito = Me.DtaConsulta.Recordset.Credito) Then
        Credito = Me.DtaConsulta.Recordset.Credito
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
 Mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones WHERE (((Transacciones.NombreCuenta)<>'**********CANCELADO*************') AND ((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ")) ORDER BY Transacciones.NumeroMovimiento"
 'Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & "))ORDER BY Transacciones.NumeroMovimiento"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
 
 '///////////Busco la Transaccion Anterior////////////
   NumeroAnterior = Me.TxtNTransacciones.Text
   Criterio = "NumeroMovimiento=" & NumeroAnterior & " "
   Me.DtaConsulta.Recordset.FindFirst Criterio
   DtaConsulta.Recordset.MovePrevious
 
   If Not DtaConsulta.Recordset.BOF Then
    NumeroTransaccion = DtaConsulta.Recordset.NumeroMovimiento
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
  Me.DBGTransacciones.Columns(5).Width = 1000
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Width = 800
  Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
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
   
    Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.NombreCuenta)<>'**********CANCELADO*************') AND ((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion "
   'Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion"
   Me.DtaTransacciones.Refresh
   If Not DtaTransacciones.Recordset.EOF Then
     Me.TxtFecha.Value = Me.DtaTransacciones.Recordset.FechaTransaccion
     Me.TxtPeriodo.Text = Me.DtaTransacciones.Recordset.Periodo
     Me.TxtNTransacciones.Text = Me.DtaTransacciones.Recordset.NumeroMovimiento
     NumeroTransaccion = Me.DtaTransacciones.Recordset.NumeroMovimiento
     Me.TxtFuente.Text = Me.DtaTransacciones.Recordset.Fuente
     
     '/////////////////////////Busco el tipo de moneda del movimiento////////////////
 Mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.TipoMoneda,IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & " ) AND ((IndiceTransaccion.NumeroMovimiento)= " & NumeroTransaccion & "))"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  If Not IsNull(Me.CmbMoneda.Text = Me.DtaConsulta.Recordset.TipoMoneda) Then
   Me.CmbMoneda.Text = Me.DtaConsulta.Recordset.TipoMoneda
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
      Me.DtaConsulta.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.CodCuentas, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, [Transacciones]![TCambio]*[Transacciones]![Debito] AS MDebito, [Transacciones]![TCambio]*[Transacciones]![Credito] AS MCredito, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito From Transacciones WHERE (((Transacciones.FechaTransaccion)=" & NumFecha1 & ") AND ((Transacciones.NumeroMovimiento)=" & NMovimiento & "))"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
      If Not IsNull(Me.DtaConsulta.Recordset.Debito) Then
       Debito = Me.DtaConsulta.Recordset.Debito
      End If
      If Not IsNull(Me.DtaConsulta.Recordset.Credito) Then
       Credito = Me.DtaConsulta.Recordset.Credito
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
  Me.DBGTransacciones.Columns(0).Button = True
    Me.DBGTransacciones.Columns(1).Locked = True
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Button = True
  Me.DBGTransacciones.Columns(6).Locked = True
  Me.DBGTransacciones.Columns(0).Width = 1500
  Me.DBGTransacciones.Columns(2).Width = 1000
  Me.DBGTransacciones.Columns(3).Caption = "Descripcion"
  Me.DBGTransacciones.Columns(4).Width = 1000
  Me.DBGTransacciones.Columns(5).Width = 1000
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Width = 800
  Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
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
 MsgBox Err.Description
  
  
  
End Sub

Private Sub CmdBorrar_Click()
On Error GoTo TipoErrs
  Dim Respuesta, Rsp
  
  
  Primero = True
  
  
  Set Rsp = DtaTransacciones.Recordset
  Respuesta = MsgBox("Esta seguro de Borrar la transaccion?", vbYesNo, "Transaccion No.: " & Me.TxtNTransacciones.Text)
   If Respuesta = 6 Then
   '//////Grabo las descripcion en los indices//////////////////////
   Me.DBGTransacciones.Enabled = True
   Mes = Month(Me.TxtFecha.Value)
   Año = Year(Me.TxtFecha.Value)
   FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
   FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
   NumFecha1 = FechaIni
   NumFecha2 = FechaFin
 
   Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & " ) AND ((IndiceTransaccion.NumeroMovimiento)= " & NumeroTransaccion & "))"
   Me.DtaConsulta.Refresh
         
       If Not DtaConsulta.Recordset.EOF Then
        
          Me.DtaConsulta.Recordset.Edit
          Me.DtaConsulta.Recordset.DescripcionMovimiento = "*****CANCELADO*****"
          Me.DtaConsulta.Recordset.Update
        
       End If
   
   
   
   
   
   Me.DtaTransacciones.Recordset.MoveFirst
    Do While Not Me.DtaTransacciones.Recordset.EOF
     Me.DtaTransacciones.Recordset.Edit
     DtaTransacciones.Recordset.NombreCuenta = "**********CANCELADO*************"
     DtaTransacciones.Recordset.DescripcionMovimiento = "**********CANCELADO*************"
     DtaTransacciones.Recordset.Debito = 0
      DtaTransacciones.Recordset.Credito = 0
     'DtaTransacciones.Recordset.Delete
     Me.DtaTransacciones.Recordset.Update
     Me.DtaTransacciones.Recordset.MoveNext
       
     Me.CmbMoneda.Enabled = False
    Loop
    Me.TxtFecha.Value = Format(Now, "dd/mm/yyyy")
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
  Me.DBGTransacciones.Columns(5).Width = 1000
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Width = 800
  Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
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

Private Sub CmdGrabar_Click()
On Error GoTo TipoErrs
 If Not Val(Me.TxtDiferencia.Text) = 0 Then
   MsgBox "El Movimiento esta Desbalanceado", vbCritical, "Sistema Contable"
   Exit Sub
  End If
Me.CmdNuevo.Enabled = True
Me.CmbMoneda.Enabled = True

'//////Grabo las descripcion en los indices//////////////////////
 Me.DBGTransacciones.Enabled = True
 Mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 If Me.CmbMoneda.Text = "" Then
   Me.CmbMoneda.Text = "Dólares"
 End If
 Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.TipoMoneda,IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & " ) AND ((IndiceTransaccion.NumeroMovimiento)= " & NumeroTransaccion & "))"
 Me.DtaConsulta.Refresh
 If Not DtaTransacciones.Recordset.EOF Then
 Me.DtaTransacciones.Recordset.MoveFirst
 End If
       
       If Not DtaConsulta.Recordset.EOF Then
        If Not Me.DBGTransacciones.Columns(3).Text = "" Then
          Me.DtaConsulta.Recordset.Edit
          Me.DtaConsulta.Recordset.TipoMoneda = Me.CmbMoneda.Text
          Me.DtaConsulta.Recordset.DescripcionMovimiento = Me.DBGTransacciones.Columns(3).Text
          Me.DtaConsulta.Recordset.Update
        End If
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

Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones Where (((Transacciones.NumeroMovimiento) = -1))"
Me.DtaTransacciones.Refresh
End If
  
  Me.DBGTransacciones.Columns(0).Button = True
    Me.DBGTransacciones.Columns(1).Locked = True
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Button = True
  Me.DBGTransacciones.Columns(6).Locked = True
  Me.DBGTransacciones.Columns(0).Width = 1500
  Me.DBGTransacciones.Columns(2).Width = 1000
  Me.DBGTransacciones.Columns(3).Caption = "Descripcion"
  Me.DBGTransacciones.Columns(4).Width = 1000
  Me.DBGTransacciones.Columns(5).Width = 1000
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Width = 800
  Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
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
  
  
TotalCredito = 0
TotalDebito = 0
Debito = 0
Credito = 0
TotalDiferencia = 0
Diferencia = 0


 
 Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub

Private Sub CmdNuevo_Click()
On Error GoTo TipoErrs
  If Not Val(Me.TxtDiferencia.Text) = 0 Then
   MsgBox "El Movimiento esta Desbalanceado", vbCritical, "Sistema Contable"
   Exit Sub
  End If
 Me.CmdNuevo.Enabled = True
 
 '//////Grabo las descripcion en los indices//////////////////////
 Me.DBGTransacciones.Enabled = True
 Mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & " ) AND ((IndiceTransaccion.NumeroMovimiento)= " & NumeroTransaccion & "))"
 Me.DtaConsulta.Refresh
 If Not DtaTransacciones.Recordset.EOF Then
 Me.DtaTransacciones.Recordset.MoveFirst
 End If
       
       If Not DtaConsulta.Recordset.EOF Then
        If Not Me.DBGTransacciones.Columns(3).Text = "" Then
          Me.DtaConsulta.Recordset.Edit
          Me.DtaConsulta.Recordset.DescripcionMovimiento = Me.DBGTransacciones.Columns(3).Text
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
 
Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones Where (((Transacciones.NumeroMovimiento) = -1))"
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
  Me.DBGTransacciones.Columns(5).Width = 1000
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Width = 800
  Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
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
  'Me.DBGTransacciones.Enabled = False

TotalCredito = 0
TotalDebito = 0
Debito = 0
Credito = 0
TotalDiferencia = 0
Diferencia = 0




 
 Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub

Private Sub CmdSalir_Click()
On Error GoTo TipoErrs
'//////Grabo las descripcion en los indices//////////////////////
 Me.DBGTransacciones.Enabled = True
 Mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & " ) AND ((IndiceTransaccion.NumeroMovimiento)= " & NumeroTransaccion & "))"
 Me.DtaConsulta.Refresh
 If Not DtaTransacciones.Recordset.EOF Then
 Me.DtaTransacciones.Recordset.MoveFirst
 End If
       If Not DtaConsulta.Recordset.EOF Then
        If Not Me.DBGTransacciones.Columns(3).Text = "" Then
          Me.DtaConsulta.Recordset.Edit
          Me.DtaConsulta.Recordset.DescripcionMovimiento = Me.DBGTransacciones.Columns(3).Text
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
On Error GoTo TipoErrs
 If Not Val(Me.TxtDiferencia.Text) = 0 Then
   MsgBox "El Movimiento esta Desbalanceado", vbCritical, "Sistema Contable"
   Exit Sub
  End If
 Primero = True
  Me.CmbMoneda.Enabled = False
  
  '//////Grabo las descripcion en los indices//////////////////////
 Me.DBGTransacciones.Enabled = True
 Mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.TipoMoneda,IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & " ) AND ((IndiceTransaccion.NumeroMovimiento)= " & NumeroTransaccion & "))"
 Me.DtaConsulta.Refresh
 If Not DtaTransacciones.Recordset.EOF Then
 Me.DtaTransacciones.Recordset.MoveFirst
 End If
       
       If Not DtaConsulta.Recordset.EOF Then
         If Not IsNull(Me.CmbMoneda.Text = Me.DtaConsulta.Recordset.TipoMoneda) Then
          'Me.CmbMoneda.Text = Me.DtaConsulta.Recordset.TipoMoneda
         End If
  
        If Not Me.DBGTransacciones.Columns(3).Text = "" Then
          Me.DtaConsulta.Recordset.Edit
          Me.DtaConsulta.Recordset.TipoMoneda = Me.CmbMoneda.Text
          Me.DtaConsulta.Recordset.DescripcionMovimiento = Me.DBGTransacciones.Columns(3).Text
          Me.DtaConsulta.Recordset.Update
        Else
          Me.DtaConsulta.Recordset.Edit
          'Me.DtaConsulta.Recordset.TipoMoneda = Me.CmbMoneda.Text
          Me.DtaConsulta.Recordset.Update
        End If
       End If
       
 TotalDiferencia = 0
 TotalCredito = 0
 TotalDebito = 0
 Debito = 0
 Credito = 0
 Diferencia = 0
 
If Me.TxtNTransacciones = 0 Then
 Mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones WHERE (((Transacciones.NombreCuenta)<>'**********CANCELADO*************') AND ((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ")) ORDER BY Transacciones.NumeroMovimiento"
 'Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & "))ORDER BY Transacciones.NumeroMovimiento"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
   '/////////Me muevo al ultimo registro/////////
   Me.DtaConsulta.Recordset.MoveLast
   NumeroTransaccion = DtaConsulta.Recordset.NumeroMovimiento
   Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.NombreCuenta)<>'**********CANCELADO*************') AND ((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion "
   'Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion"
   Me.DtaTransacciones.Refresh
   
   If Not DtaTransacciones.Recordset.EOF Then
     Me.TxtFecha.Value = Me.DtaTransacciones.Recordset.FechaTransaccion
     Me.TxtPeriodo.Text = Me.DtaTransacciones.Recordset.Periodo
     Me.TxtNTransacciones.Text = Me.DtaTransacciones.Recordset.NumeroMovimiento
     NumeroTransaccion = Me.DtaTransacciones.Recordset.NumeroMovimiento
     Me.TxtFuente.Text = Me.DtaTransacciones.Recordset.Fuente
         '/////////////////////////Busco el tipo de moneda del movimiento////////////////
 Mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.TipoMoneda,IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & " ) AND ((IndiceTransaccion.NumeroMovimiento)= " & NumeroTransaccion & "))"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  If Not IsNull(Me.DtaConsulta.Recordset.TipoMoneda) Then
   Me.CmbMoneda.Text = Me.DtaConsulta.Recordset.TipoMoneda
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
      Me.DtaConsulta.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.CodCuentas, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, [Transacciones]![TCambio]*[Transacciones]![Debito] AS MDebito, [Transacciones]![TCambio]*[Transacciones]![Credito] AS MCredito, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito From Transacciones WHERE (((Transacciones.FechaTransaccion)=" & NumFecha1 & ") AND ((Transacciones.NumeroMovimiento)=" & NMovimiento & "))"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
       If Not IsNull(Me.DtaConsulta.Recordset.Debito) Then
       Debito = Me.DtaConsulta.Recordset.Debito
       End If
       If Not IsNull(Credito = Me.DtaConsulta.Recordset.Credito) Then
        Credito = Me.DtaConsulta.Recordset.Credito
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
 Mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones WHERE (((Transacciones.NombreCuenta)<>'**********CANCELADO*************') AND ((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ")) ORDER BY Transacciones.NumeroMovimiento"
 'Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,  Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & "))ORDER BY Transacciones.NumeroMovimiento"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
 
 '///////////Busco la Transaccion Siguiente////////////
   NumeroAnterior = Me.TxtNTransacciones.Text
   Criterio = "NumeroMovimiento=" & NumeroAnterior & " "
   Me.DtaConsulta.Recordset.FindLast Criterio
   DtaConsulta.Recordset.MoveNext
   
 If Not DtaConsulta.Recordset.EOF Then
   NumeroTransaccion = DtaConsulta.Recordset.NumeroMovimiento
 Else '//En caso que no se encuentre ninguna transaccion
  MsgBox "Esta es la ultima Transaccion del Periodo", vbInformation, "Sistema Contable"
    Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.NombreCuenta)<>'**********CANCELADO*************') AND ((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion "
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
  Me.DBGTransacciones.Columns(5).Width = 1000
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Width = 800
  Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
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
   Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.NombreCuenta)<>'**********CANCELADO*************') AND ((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion "
   'Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,  Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion"
   Me.DtaTransacciones.Refresh
   If Not DtaTransacciones.Recordset.EOF Then
     Me.TxtFecha.Value = Me.DtaTransacciones.Recordset.FechaTransaccion
     Me.TxtPeriodo.Text = Me.DtaTransacciones.Recordset.Periodo
     Me.TxtNTransacciones.Text = Me.DtaTransacciones.Recordset.NumeroMovimiento
      NumeroTransaccion = Me.DtaTransacciones.Recordset.NumeroMovimiento
     Me.TxtFuente.Text = Me.DtaTransacciones.Recordset.Fuente
     
     '/////////////////////////Busco el tipo de moneda del movimiento////////////////
 Mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.TipoMoneda,IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & " ) AND ((IndiceTransaccion.NumeroMovimiento)= " & NumeroTransaccion & "))"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  If Not IsNull(Me.DtaConsulta.Recordset.TipoMoneda) Then
   Me.CmbMoneda.Text = Me.DtaConsulta.Recordset.TipoMoneda
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
      Me.DtaConsulta.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.CodCuentas, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, [Transacciones]![TCambio]*[Transacciones]![Debito] AS MDebito, [Transacciones]![TCambio]*[Transacciones]![Credito] AS MCredito, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito From Transacciones WHERE (((Transacciones.FechaTransaccion)=" & NumFecha1 & ") AND ((Transacciones.NumeroMovimiento)=" & NMovimiento & "))"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
      If Not IsNull(Me.DtaConsulta.Recordset.Debito) Then
       Debito = Me.DtaConsulta.Recordset.Debito
      End If
      If Not IsNull(Me.DtaConsulta.Recordset.Credito) Then
       Credito = Me.DtaConsulta.Recordset.Credito
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
     Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,  Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroAnterior & ")) ORDER BY Transacciones.NTransaccion"
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


  
   Me.DBGTransacciones.Columns(0).Button = True
     Me.DBGTransacciones.Columns(1).Locked = True
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Button = True
  Me.DBGTransacciones.Columns(6).Locked = True
  Me.DBGTransacciones.Columns(0).Width = 1500
  Me.DBGTransacciones.Columns(2).Width = 1000
  Me.DBGTransacciones.Columns(3).Caption = "Descripcion"
  Me.DBGTransacciones.Columns(4).Width = 1000
  Me.DBGTransacciones.Columns(5).Width = 1000
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Width = 800
  Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
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
 MsgBox Err.Description
End Sub

Private Sub DBGTransacciones_AfterColEdit(ByVal ColIndex As Integer)
On Error GoTo TipoErrs
Dim Descripcion As String, TipoCunta As String, Numero As String, Fecha As Long
Dim MontoTasa As Double
'Este Procedimiento es solo cuando se ejecuta directamente de Recepcion
QueProducto = "Transacciones"
Me.CmbMoneda.Enabled = False

'/////Busco cambios en las claves del movimiento///////////


Select Case ColIndex
  Case 0
    '////////////Verifico la cuenta///////////////
  
  
       Criterio = "CodCuentas='" & Me.DBGTransacciones.Columns(0).Text & "'"
       Me.DtaCuentas.Recordset.FindFirst Criterio
       If Not DtaCuentas.Recordset.NoMatch Then
         TipoCuenta = DtaCuentas.Recordset.TipoCuenta
         TipoMoneda = DtaCuentas.Recordset.TipoMoneda

         Select Case TipoMoneda
            Case "Córdobas"
                      Fecha = Me.TxtFecha.Value
                      Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      Me.DtaTasas.Refresh
                If Not DtaTasas.Recordset.EOF Then
                 MontoTasa = Me.DtaTasas.Recordset.MontoCordobas
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
             Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
             Me.DtaTasas.Refresh
             If Not DtaTasas.Recordset.EOF Then
                MontoTasa = Me.DtaTasas.Recordset.MontoCordobas
                If MontoTasa = 0 Then
                  MontoTasa = 1
                End If
            
               Select Case Me.CmbMoneda.Text
                  Case "Córdobas"
                    Me.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                  Case "Dólares"
                    Me.DBGTransacciones.Columns(7).Text = 1
                  Case "Libras"
                    MontoTasa = Me.DtaTasas.Recordset.MontoLibras
                    Me.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)

                    
                 End Select
                Else
                  Me.DBGTransacciones.Columns(7).Text = 1
               End If
            
            Case "Libras"
                      Fecha = Me.TxtFecha.Value
                      Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      Me.DtaTasas.Refresh
                If Not DtaTasas.Recordset.EOF Then
                MontoTasa = Me.DtaTasas.Recordset.MontoLibras
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
         
   TipoCuenta = Me.DtaCuentas.Recordset.TipoCuenta
   CodigoCuenta = DtaCuentas.Recordset.CodCuentas
  If TipoCuenta = "Cuentas de Banco Cordobas" Or TipoCuenta = "Cuentas de Banco Dolares" Or TipoCuenta = "Cuentas de Efectivo" Then

   If Primero = True Then
     Primero = False
        Me.DtaConsulta.RecordSource = "SELECT NConsecutivoVoucher.CodCuenta, NConsecutivoVoucher.ConsecutivoVoucher, NConsecutivoVoucher.NPeriodo From NConsecutivoVoucher Where (((NConsecutivoVoucher.CodCuenta) = '" & CodigoCuenta & "') And ((NConsecutivoVoucher.NPeriodo) = " & NumeroPeriodo & "))"
        Me.DtaConsulta.Refresh
        If DtaConsulta.Recordset.EOF Then
           Me.DtaConsulta.Recordset.AddNew
             Me.DtaConsulta.Recordset.CodCuenta = CodigoCuenta
             Me.DtaConsulta.Recordset.NPeriodo = NumeroPeriodo
             Me.DtaConsulta.Recordset.ConsecutivoVoucher = 1
           Me.DtaConsulta.Recordset.Update
           NumeroVoucher = 1
        Else
           Me.DtaConsulta.Recordset.Edit
            Me.DtaConsulta.Recordset.ConsecutivoVoucher = Me.DtaConsulta.Recordset.ConsecutivoVoucher + 1
           Me.DtaConsulta.Recordset.Update
         NumeroVoucher = Me.DtaConsulta.Recordset.ConsecutivoVoucher
        End If
     Else
       ' FrmCheque.DtaTransacciones.Recordset.MoveLast
     
     End If
        ConsecutivoVoucher = Month(FrmTransacciones.TxtFecha.Value)
        Select Case TipoCuenta
           Case "Cuentas de Banco Cordobas"
              Numero = "BC " & NumeroVoucher & "/" & ConsecutivoVoucher
             
           Case "Cuentas de Banco Dolares"
              Numero = "BD " & NumeroVoucher & "/" & ConsecutivoVoucher
           Case "Cuentas de Efectivo"
              Numero = "CASH " & NumeroVoucher & "/" & ConsecutivoVoucher
        
         End Select
        
     End If

         
         
         Me.DBGTransacciones.Columns(2).Text = Numero
         Me.DBGTransacciones.Columns(1).Text = DtaCuentas.Recordset.DescripcionCuentas
         Me.DBGTransacciones.Columns(10).Text = Me.TxtFecha.Value
         Me.DBGTransacciones.Columns(11).Text = NumeroPeriodo
         Me.DBGTransacciones.Columns(13).Text = Me.TxtFuente.Text
         Me.DBGTransacciones.Columns(14).Text = Me.TxtFecha.Value
         Me.DBGTransacciones.Columns(15).Text = NumeroTransaccion
         Me.DBGTransacciones.Columns(6).Text = "Debito"
         Me.DBGTransacciones.Columns(9).Locked = True
         Me.DBGTransacciones.Columns(9).Locked = True
         Me.DBGTransacciones.Columns(8).Locked = False

       Else
               
         MsgBox "La cuenta digitada no es correcta", vbCritical, "Sistema Contable"
         NumeroTransaccion = Me.TxtNTransacciones.Text
         Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,  Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion"
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
  Me.DBGTransacciones.Columns(5).Width = 1000
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Width = 800
  Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
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
    Me.DtaCuentas.Recordset.FindFirst Criterio
    If Not DtaCuentas.Recordset.NoMatch Then
     TipoCuenta = Me.DtaCuentas.Recordset.TipoCuenta
     If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Cuentas de Efectivo" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Cuentas x Pagar" Or TipoCuenta = "Cuentas de Gastos" Or TipoCuenta = "Cuentas de Banco Cordobas" Or TipoCuenta = "Cuentas de Banco Dolares" Then
      
   '//////Busco si tiene saldo en el historial del perido actual
      Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
      Me.DtaHistorial.Refresh
       If DtaHistorial.Recordset.EOF Then
        '////Si no existe registro para este Periodo Busco los ///
        '////saldos del periodo anterior para que sean incial////
         Criterio = "NPeriodo=" & NumeroPeriodo & " "
         Me.DtaPeriodos.Recordset.FindFirst Criterio
        If Not DtaPeriodos.Recordset.NoMatch Then
'////Busco el numero del periodo anterior para hacer la consulta
          DtaPeriodos.Recordset.MovePrevious
          NumeroPeriodoAnterior = DtaPeriodos.Recordset.NPeriodo
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodoAnterior & "))"
          Me.DtaHistorial.Refresh
        
          If DtaHistorial.Recordset.EOF Then
       '//////si el periodo anterior no tiene saldo////
       '/////Y el periodo actual tampoco tiene saldo/////////
           Me.DtaHistorial.Recordset.AddNew
             Me.DtaHistorial.Recordset.CodCuenta = CodigoCuenta
             Me.DtaHistorial.Recordset.NPeriodo = NumeroPeriodo
             Me.DtaHistorial.Recordset.SaldoInicial = 0
             Me.DtaHistorial.Recordset.SaldoFinal = Val(Me.DtaHistorial.Recordset.SaldoInicial) + Diferencia
            Me.DtaHistorial.Recordset.Update
         Else
      '//////Si el periodo anterior tiene saldo////////
      '/////y el periodo actual no tiene saldo//////////
         SaldoFinal = Me.DtaHistorial.Recordset.SaldoFinal
            Me.DtaHistorial.Recordset.AddNew
             Me.DtaHistorial.Recordset.CodCuenta = CodigoCuenta
             Me.DtaHistorial.Recordset.NPeriodo = NumeroPeriodo
             Me.DtaHistorial.Recordset.SaldoInicial = SaldoFinal
             Me.DtaHistorial.Recordset.SaldoFinal = Val(Me.DtaHistorial.Recordset.SaldoInicial) + Diferencia
            Me.DtaHistorial.Recordset.Update
         
         End If
        End If
       Else '////////Si la cuenta tiene saldo en el periodo actual
        '////Si existe registro para este Periodo Busco los ///
        '////saldos del periodo anterior para que sean incial////
         Criterio = "NPeriodo=" & NumeroPeriodo & " "
         Me.DtaPeriodos.Recordset.FindFirst Criterio
        If Not DtaPeriodos.Recordset.NoMatch Then
'////Busco el numero del periodo anterior para hacer la consulta
          DtaPeriodos.Recordset.MovePrevious
          NumeroPeriodoAnterior = DtaPeriodos.Recordset.NPeriodo
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodoAnterior & "))"
          Me.DtaHistorial.Refresh
        
          If DtaHistorial.Recordset.EOF Then
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
          Me.DtaHistorial.Refresh
          
       '//////si el periodo anterior no tiene saldo////
       '/////Y el periodo actual tiene saldo/////////
           Me.DtaHistorial.Recordset.Edit
             Me.DtaHistorial.Recordset.SaldoInicial = 0
             Me.DtaHistorial.Recordset.SaldoFinal = Val(Me.DtaHistorial.Recordset.SaldoFinal) + Diferencia
            Me.DtaHistorial.Recordset.Update
         Else
          SaldoFinal = Me.DtaHistorial.Recordset.SaldoFinal
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
          Me.DtaHistorial.Refresh
      '//////Si el periodo anterior tiene saldo////////
      '/////y el periodo actual tiene saldo//////////
        
            Me.DtaHistorial.Recordset.Edit
             Me.DtaHistorial.Recordset.SaldoInicial = SaldoFinal
             Me.DtaHistorial.Recordset.SaldoFinal = Val(Me.DtaHistorial.Recordset.SaldoFinal) + Diferencia
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
         Me.DtaPeriodos.Recordset.FindFirst Criterio
        If Not DtaPeriodos.Recordset.NoMatch Then
'////Busco el numero del periodo anterior para hacer la consulta
          DtaPeriodos.Recordset.MovePrevious
          NumeroPeriodoAnterior = DtaPeriodos.Recordset.NPeriodo
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodoAnterior & "))"
          Me.DtaHistorial.Refresh
        
          If DtaHistorial.Recordset.EOF Then
       '//////si el periodo anterior no tiene saldo////
       '/////Y el periodo actual tampoco tiene saldo/////////
           Me.DtaHistorial.Recordset.AddNew
             Me.DtaHistorial.Recordset.CodCuenta = CodigoCuenta
             Me.DtaHistorial.Recordset.NPeriodo = NumeroPeriodo
             Me.DtaHistorial.Recordset.SaldoInicial = 0
             Me.DtaHistorial.Recordset.SaldoFinal = -Diferencia
            Me.DtaHistorial.Recordset.Update
         Else
      '//////Si el periodo anterior tiene saldo////////
      '/////y el periodo actual no tiene saldo//////////
         SaldoFinal = Me.DtaHistorial.Recordset.SaldoFinal
            Me.DtaHistorial.Recordset.AddNew
             Me.DtaHistorial.Recordset.CodCuenta = CodigoCuenta
             Me.DtaHistorial.Recordset.NPeriodo = NumeroPeriodo
             Me.DtaHistorial.Recordset.SaldoInicial = SaldoFinal
             Me.DtaHistorial.Recordset.SaldoFinal = -Diferencia
            Me.DtaHistorial.Recordset.Update
         
         End If
        End If
       Else '////////Si la cuenta tiene saldo en el periodo actual
        '////Si existe registro para este Periodo Busco los ///
        '////saldos del periodo anterior para que sean incial////
         Criterio = "NPeriodo=" & NumeroPeriodo & " "
         Me.DtaPeriodos.Recordset.FindFirst Criterio
        If Not DtaPeriodos.Recordset.NoMatch Then
'////Busco el numero del periodo anterior para hacer la consulta
          DtaPeriodos.Recordset.MovePrevious
          NumeroPeriodoAnterior = DtaPeriodos.Recordset.NPeriodo
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodoAnterior & "))"
          Me.DtaHistorial.Refresh
        
          If DtaHistorial.Recordset.EOF Then
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
          Me.DtaHistorial.Refresh
          
       '//////si el periodo anterior no tiene saldo////
       '/////Y el periodo actual tiene saldo/////////
           Me.DtaHistorial.Recordset.Edit
             Me.DtaHistorial.Recordset.SaldoInicial = 0
             Me.DtaHistorial.Recordset.SaldoFinal = Val(Me.DtaHistorial.Recordset.SaldoFinal) - Diferencia
            Me.DtaHistorial.Recordset.Update
         Else
          SaldoFinal = Me.DtaHistorial.Recordset.SaldoFinal
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
          Me.DtaHistorial.Refresh
      '//////Si el periodo anterior tiene saldo////////
      '/////y el periodo actual tiene saldo//////////
        
            Me.DtaHistorial.Recordset.Edit
             Me.DtaHistorial.Recordset.SaldoInicial = SaldoFinal
             Me.DtaHistorial.Recordset.SaldoFinal = Val(Me.DtaHistorial.Recordset.SaldoFinal) - Diferencia
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
    Me.DtaCuentas.Recordset.FindFirst Criterio
    If Not DtaCuentas.Recordset.NoMatch Then
     TipoCuenta = Me.DtaCuentas.Recordset.TipoCuenta
     If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Cuentas de Efectivo" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Cuentas x Pagar" Or TipoCuenta = "Cuentas de Gastos" Or TipoCuenta = "Cuentas de Banco Cordobas" Or TipoCuenta = "Cuentas de Banco Dolares" Then
      
   '//////Busco si tiene saldo en el historial del perido actual
      Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
      Me.DtaHistorial.Refresh
       If DtaHistorial.Recordset.EOF Then
        '////Si no existe registro para este Periodo Busco los ///
        '////saldos del periodo anterior para que sean incial////
         Criterio = "NPeriodo=" & NumeroPeriodo & " "
         Me.DtaPeriodos.Recordset.FindFirst Criterio
        If Not DtaPeriodos.Recordset.NoMatch Then
'////Busco el numero del periodo anterior para hacer la consulta
          DtaPeriodos.Recordset.MovePrevious
          NumeroPeriodoAnterior = DtaPeriodos.Recordset.NPeriodo
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodoAnterior & "))"
          Me.DtaHistorial.Refresh
        
          If DtaHistorial.Recordset.EOF Then
       '//////si el periodo anterior no tiene saldo////
       '/////Y el periodo actual tampoco tiene saldo/////////
           Me.DtaHistorial.Recordset.AddNew
             Me.DtaHistorial.Recordset.CodCuenta = CodigoCuenta
             Me.DtaHistorial.Recordset.NPeriodo = NumeroPeriodo
             Me.DtaHistorial.Recordset.SaldoInicial = 0
             Me.DtaHistorial.Recordset.SaldoFinal = -Diferencia
            Me.DtaHistorial.Recordset.Update
         Else
      '//////Si el periodo anterior tiene saldo////////
      '/////y el periodo actual no tiene saldo//////////
         SaldoFinal = Me.DtaHistorial.Recordset.SaldoFinal
            Me.DtaHistorial.Recordset.AddNew
             Me.DtaHistorial.Recordset.CodCuenta = CodigoCuenta
             Me.DtaHistorial.Recordset.NPeriodo = NumeroPeriodo
             Me.DtaHistorial.Recordset.SaldoInicial = SaldoFinal
             Me.DtaHistorial.Recordset.SaldoFinal = -Diferencia
            Me.DtaHistorial.Recordset.Update
         
         End If
        End If
       Else '////////Si la cuenta tiene saldo en el periodo actual
        '////Si existe registro para este Periodo Busco los ///
        '////saldos del periodo anterior para que sean incial////
         Criterio = "NPeriodo=" & NumeroPeriodo & " "
         Me.DtaPeriodos.Recordset.FindFirst Criterio
        If Not DtaPeriodos.Recordset.NoMatch Then
'////Busco el numero del periodo anterior para hacer la consulta
          DtaPeriodos.Recordset.MovePrevious
          NumeroPeriodoAnterior = DtaPeriodos.Recordset.NPeriodo
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodoAnterior & "))"
          Me.DtaHistorial.Refresh
        
          If DtaHistorial.Recordset.EOF Then
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
          Me.DtaHistorial.Refresh
          
       '//////si el periodo anterior no tiene saldo////
       '/////Y el periodo actual tiene saldo/////////
           Me.DtaHistorial.Recordset.Edit
             Me.DtaHistorial.Recordset.SaldoInicial = 0
             Me.DtaHistorial.Recordset.SaldoFinal = Val(Me.DtaHistorial.Recordset.SaldoFinal) - Diferencia
            Me.DtaHistorial.Recordset.Update
         Else
          SaldoFinal = Me.DtaHistorial.Recordset.SaldoFinal
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
          Me.DtaHistorial.Refresh
      '//////Si el periodo anterior tiene saldo////////
      '/////y el periodo actual tiene saldo//////////
        
            Me.DtaHistorial.Recordset.Edit
             Me.DtaHistorial.Recordset.SaldoInicial = SaldoFinal
             Me.DtaHistorial.Recordset.SaldoFinal = Val(Me.DtaHistorial.Recordset.SaldoFinal) - Diferencia
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
         Me.DtaPeriodos.Recordset.FindFirst Criterio
        If Not DtaPeriodos.Recordset.NoMatch Then
'////Busco el numero del periodo anterior para hacer la consulta
          DtaPeriodos.Recordset.MovePrevious
          NumeroPeriodoAnterior = DtaPeriodos.Recordset.NPeriodo
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodoAnterior & "))"
          Me.DtaHistorial.Refresh
        
          If DtaHistorial.Recordset.EOF Then
       '//////si el periodo anterior no tiene saldo////
       '/////Y el periodo actual tampoco tiene saldo/////////
           Me.DtaHistorial.Recordset.AddNew
             Me.DtaHistorial.Recordset.CodCuenta = CodigoCuenta
             Me.DtaHistorial.Recordset.NPeriodo = NumeroPeriodo
             Me.DtaHistorial.Recordset.SaldoInicial = 0
             Me.DtaHistorial.Recordset.SaldoFinal = Diferencia
            Me.DtaHistorial.Recordset.Update
         Else
      '//////Si el periodo anterior tiene saldo////////
      '/////y el periodo actual no tiene saldo//////////
         SaldoFinal = Me.DtaHistorial.Recordset.SaldoFinal
            Me.DtaHistorial.Recordset.AddNew
             Me.DtaHistorial.Recordset.CodCuenta = CodigoCuenta
             Me.DtaHistorial.Recordset.NPeriodo = NumeroPeriodo
             Me.DtaHistorial.Recordset.SaldoInicial = SaldoFinal
             Me.DtaHistorial.Recordset.SaldoFinal = Diferencia
            Me.DtaHistorial.Recordset.Update
         
         End If
        End If
       Else '////////Si la cuenta tiene saldo en el periodo actual
        '////Si existe registro para este Periodo Busco los ///
        '////saldos del periodo anterior para que sean incial////
         Criterio = "NPeriodo=" & NumeroPeriodo & " "
         Me.DtaPeriodos.Recordset.FindFirst Criterio
        If Not DtaPeriodos.Recordset.NoMatch Then
'////Busco el numero del periodo anterior para hacer la consulta
          DtaPeriodos.Recordset.MovePrevious
          NumeroPeriodoAnterior = DtaPeriodos.Recordset.NPeriodo
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodoAnterior & "))"
          Me.DtaHistorial.Refresh
        
          If DtaHistorial.Recordset.EOF Then
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
          Me.DtaHistorial.Refresh
          
       '//////si el periodo anterior no tiene saldo////
       '/////Y el periodo actual tiene saldo/////////
           Me.DtaHistorial.Recordset.Edit
             Me.DtaHistorial.Recordset.SaldoInicial = 0
             Me.DtaHistorial.Recordset.SaldoFinal = Val(Me.DtaHistorial.Recordset.SaldoFinal) + Diferencia
            Me.DtaHistorial.Recordset.Update
         Else
          SaldoFinal = Me.DtaHistorial.Recordset.SaldoFinal
          Me.DtaHistorial.RecordSource = "SELECT Historial.CodCuenta, Historial.NPeriodo, Historial.SaldoInicial, Historial.SaldoFinal From Historial WHERE (((Historial.CodCuenta)='" & CodigoCuenta & "') AND ((Historial.NPeriodo)=" & NumeroPeriodo & "))"
          Me.DtaHistorial.Refresh
      '//////Si el periodo anterior tiene saldo////////
      '/////y el periodo actual tiene saldo//////////
        
            Me.DtaHistorial.Recordset.Edit
             Me.DtaHistorial.Recordset.SaldoInicial = SaldoFinal
             Me.DtaHistorial.Recordset.SaldoFinal = Val(Me.DtaHistorial.Recordset.SaldoFinal) + Diferencia
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
   Select Case ColIndex
    Case 0
      Mes = Month(Me.TxtFecha.Value)
      Año = Year(Me.TxtFecha.Value)
      FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
      FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
      NumFecha1 = FechaIni
      NumFecha2 = FechaFin
 
      Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
      Me.DtaConsulta.Refresh
      If Not DtaConsulta.Recordset.EOF Then
        Me.TxtPeriodo.Text = DtaConsulta.Recordset.Periodo
        NumeroPeriodo = DtaConsulta.Recordset.NPeriodo
        If Val(Me.TxtNTransacciones.Text) = 0 Then
         NumeroTransaccion = DtaConsulta.Recordset.NTransacciones
        End If
        EstadoPeriodo = DtaConsulta.Recordset.EstadoPeriodo
      
      '////////////Edito los datos del Periodo///////////
     If Val(Me.TxtNTransacciones.Text) = 0 Then
     
     
     
     
      Me.DtaConsulta.Recordset.Edit
        DtaConsulta.Recordset.NTransacciones = DtaConsulta.Recordset.NTransacciones + 1
      Me.DtaConsulta.Recordset.Update
      NumeroTransaccion = DtaConsulta.Recordset.NTransacciones
      FrmTransacciones.TxtNTransacciones.Text = NumeroTransaccion
      '////////Edito los Datos de los indices de Transacciones//////
         
          Me.DtaIndice.Recordset.AddNew
          Me.DtaIndice.Recordset.FechaTransaccion = Me.TxtFecha.Value
          Me.DtaIndice.Recordset.NumeroMovimiento = NumeroTransaccion
          Me.DtaIndice.Recordset.DescripcionMovimiento = Me.DBGTransacciones.Columns(1).Text
          Me.DtaIndice.Recordset.Fuente = Me.TxtFuente.Text
          Me.DtaIndice.Recordset.NPeriodo = NumeroPeriodo
          Me.DtaIndice.Recordset.Update
      
      
      
     
     
     
     End If
   End If
  End Select
  
  Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub DBGTransacciones_AfterUpdate()

 Debito = 0
 Credito = 0
 TotalDebito = 0
 TotalCredito = 0
      NumFecha1 = Me.TxtFecha.Value
      NMovimiento = Val(Me.TxtNTransacciones)
      Me.DtaConsulta.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.CodCuentas, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, [Transacciones]![TCambio]*[Transacciones]![Debito] AS MDebito, [Transacciones]![TCambio]*[Transacciones]![Credito] AS MCredito, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito From Transacciones WHERE (((Transacciones.FechaTransaccion)=" & NumFecha1 & ") AND ((Transacciones.NumeroMovimiento)=" & NMovimiento & "))"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
       Debito = Me.DtaConsulta.Recordset.Debito
       Credito = Me.DtaConsulta.Recordset.Credito
       TotalDebito = TotalDebito + Debito
       TotalCredito = TotalCredito + Credito
       DtaConsulta.Recordset.MoveNext
      Loop
Me.DBGTransacciones.Columns(7).NumberFormat = "##,##0.000000"
Me.TxtCredito.Text = Format(TotalCredito, "##,##0.00")
Me.TxtDebito.Text = Format(TotalDebito, "##,##0.00")
Me.TxtDiferencia.Text = Format(TotalDebito - TotalCredito, "##,##0.00")

Me.CmdNuevo.Enabled = False
End Sub

Private Sub DBGTransacciones_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
On Error GoTo TipoErrs
If ColIndex = 8 Or ColIndex = 9 Then
 If Me.DBGTransacciones.Columns(6).Text = "Debito" Then
      Me.DBGTransacciones.Columns(9).Locked = True
      Me.DBGTransacciones.Columns(8).Locked = False
  ElseIf Me.DBGTransacciones.Columns(6).Text = "Credito" Then
  
     Me.DBGTransacciones.Columns(9).Locked = False
     Me.DBGTransacciones.Columns(8).Locked = True
 
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

Private Sub DBGTransacciones_BeforeUpdate(Cancel As Integer)
On Error GoTo TipoErrs
 If Me.DBGTransacciones.Columns(6).Text = "" Then
   Me.DBGTransacciones.Columns(6).Text = "Debito"
 End If
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub DBGTransacciones_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo TipoErrs
Select Case ColIndex
  Case 0
  
  QueProducto = "Transacciones"
  FrmConsulta.Show 1
  Case 6
    Set c = DBGTransacciones.Columns(ColIndex)
      With List1
      .Left = Me.DBGTransacciones.Left + c.Left
      .Top = DBGTransacciones.Top + DBGTransacciones.RowTop(DBGTransacciones.Row) + DBGTransacciones.RowHeight
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



Private Sub DBGTransacciones_GotFocus()
On Error GoTo TipoErrs
 Mes = Month(Me.TxtFecha.Value)
      Año = Year(Me.TxtFecha.Value)
      FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
      FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
      NumFecha1 = FechaIni
      NumFecha2 = FechaFin
 
      Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
      Me.DtaConsulta.Refresh
      If Not DtaConsulta.Recordset.EOF Then
        Me.TxtPeriodo.Text = DtaConsulta.Recordset.Periodo
        NumeroPeriodo = DtaConsulta.Recordset.NPeriodo
        If Val(Me.TxtNTransacciones.Text) = 0 Then
        NumeroTransaccion = DtaConsulta.Recordset.NTransacciones
        End If
        EstadoPeriodo = DtaConsulta.Recordset.EstadoPeriodo
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
On Error GoTo TipoErrs
 If KeyCode = 113 Then
  QueProducto = "Transacciones"
  FrmConsulta.Show 1
 End If
 
 If KeyCode = 114 Then
  Indice = 1
     
  Criterio = "CodCuentas='" & Me.DBGTransacciones.Columns(0).Text & "'"
  Me.DtaCuentas.Recordset.FindFirst Criterio
  If Not DtaCuentas.Recordset.NoMatch Then
     TipoMoneda = DtaCuentas.Recordset.TipoMoneda
  End If
   FrmConvertir.LblNombre.Caption = "Monto " & TipoMoneda
   FrmConvertir.TxtTasa.Text = Me.DBGTransacciones.Columns(7).Text
   
   FrmConvertir.Show 1
  
 End If
 

 
Exit Sub
TipoErrs:
 ControlErrores
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
MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo TipoErrs
Primero = True
With Me.DtaNacceso
   .DatabaseName = Ruta
   .Connect = Conexion
End With

With Me.DtaPeriodos
   .DatabaseName = Ruta
   .Connect = Conexion
End With

With Me.DtaHistorial
   .DatabaseName = Ruta
   .Connect = Conexion
End With

With Me.DtaIndice
   .DatabaseName = Ruta
   .Connect = Conexion
End With

With Me.DtaTransaccionesNuevas
   .DatabaseName = Ruta
   .Connect = Conexion
End With

With Me.DtaCuentas
   .DatabaseName = Ruta
   .Connect = Conexion
End With


With Me.DtaTasas
   .DatabaseName = Ruta
   .Connect = Conexion
End With

With Me.DtaConsulta
   .DatabaseName = Ruta
   .Connect = Conexion
End With

With Me.DtaTransacciones
   .DatabaseName = Ruta
   .Connect = Conexion
End With

Me.TxtFecha.Value = Format(Now, "dd/mm/yyyy")

Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones Where (((Transacciones.NumeroMovimiento) = -1))"
Me.DtaTransacciones.Refresh

  Me.CmbMoneda.Text = "Dólares"
  
  
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
  Me.DBGTransacciones.Columns(5).Width = 1000
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Width = 800
  Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
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
  Me.DBGTransacciones.Enabled = False
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub Text2_Change()

End Sub

Private Sub TDBGrid1_Click()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo TipoErrs
  If Not Val(Me.TxtDiferencia.Text) = 0 Then
   MsgBox "El Movimiento esta Desbalanceado", vbCritical, "Sistema Contable"
   Cancel = 1
  End If
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub List1_DblClick()
Me.DBGTransacciones.Columns(6).Text = Me.List1.Text
 Select Case List1.Text
   Case "Debito"
      Me.DBGTransacciones.Columns(9).Locked = True
      Me.DBGTransacciones.Columns(8).Locked = False
      
   Case "Credito"
     Me.DBGTransacciones.Columns(9).Locked = False
     Me.DBGTransacciones.Columns(8).Locked = True
   
 End Select
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

Private Sub SmartButton1_Click()

On Error GoTo TipoErrs
  Dim Respuesta, Rsp
  
  If Not Me.DBGTransacciones.Columns(8).Text = "0.00" Then
    MsgBox "Debe llenar de Cero el campo del Debito"
    Exit Sub
  End If
  
   If Not Me.DBGTransacciones.Columns(9).Text = "0.00" Then
    MsgBox "Debe llenar de Cero el campo del Credito"
    Exit Sub
  End If
  
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
     DtaTransacciones.Recordset.Delete
    
      
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
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Button = True
  Me.DBGTransacciones.Columns(6).Locked = True
  Me.DBGTransacciones.Columns(0).Width = 1500
  Me.DBGTransacciones.Columns(2).Width = 1000
  Me.DBGTransacciones.Columns(3).Caption = "Descripcion"
  Me.DBGTransacciones.Columns(4).Width = 1000
  Me.DBGTransacciones.Columns(5).Width = 1000
  Me.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
  Me.DBGTransacciones.Columns(6).Width = 800
  Me.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
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
TipoErrs:
 ControlErrores
End Sub

Private Sub TxtFecha_GotFocus()
On Error GoTo TipoErrs
 Me.DBGTransacciones.Enabled = True
 Mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  Me.TxtPeriodo.Text = DtaConsulta.Recordset.Periodo
  NumeroPeriodo = DtaConsulta.Recordset.NPeriodo
  NumeroTransaccion = DtaConsulta.Recordset.NTransacciones
  EstadoPeriodo = DtaConsulta.Recordset.EstadoPeriodo
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
DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) = " & NumFecha & "))ORDER BY Tasas.FechaTasas"
DtaTasas.Refresh

If Not DtaTasas.Recordset.EOF Then
Fecha = Format(DtaTasas.Recordset.FechaTasas, "dd/mm/yyyy")
   
    Encontrado = True
    Cambio = DtaTasas.Recordset.MontoCordobas
    MDIPrimero.StatusBar2.Panels(2) = "Tasa Dolar: " & Format(Cambio, "##,##0.00") & "    " & "Tasa Libras: " & Format(DtaTasas.Recordset.MontoLibras, "##,##0.00")
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
Dim NumFecha As Long
Mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  Me.TxtPeriodo.Text = DtaConsulta.Recordset.Periodo
 
 End If
 
 '///////Verifico si esta registrada la fecha de la tasa//////

NumFecha = Me.TxtFecha.Value
DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) = " & NumFecha & "))ORDER BY Tasas.FechaTasas"
DtaTasas.Refresh

If Not DtaTasas.Recordset.EOF Then
Fecha = Format(DtaTasas.Recordset.FechaTasas, "dd/mm/yyyy")
   
    Encontrado = True
    Cambio = DtaTasas.Recordset.MontoCordobas
    MDIPrimero.StatusBar2.Panels(2) = "Tasa Dolar: " & Format(Cambio, "##,##0.00") & "    " & "Tasa Libras: " & Format(DtaTasas.Recordset.MontoLibras, "##,##0.00")
End If

If Not Encontrado Then
  MsgBox "La tasa de esta Fecha no ha sido Grabada"
  Tasa = False
  frmTasa2.Show 1
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
Mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  Me.TxtPeriodo.Text = DtaConsulta.Recordset.Periodo
  NumeroPeriodo = DtaConsulta.Recordset.NPeriodo
  NumeroTransaccion = DtaConsulta.Recordset.NTransacciones
  EstadoPeriodo = DtaConsulta.Recordset.EstadoPeriodo
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
Fecha = Format(DtaTasas.Recordset.FechaTasas, "dd/mm/yyyy")
   
    Encontrado = True
    Cambio = DtaTasas.Recordset.MontoCordobas
    MDIPrimero.StatusBar2.Panels(2) = "Tasa Dolar: " & Format(Cambio, "##,##0.00") & "    " & "Tasa Libras: " & Format(DtaTasas.Recordset.MontoLibras, "##,##0.00")
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

Private Sub TxtNTransacciones_LostFocus()
On Error GoTo TipoErrs
 Mes = Month(Me.TxtFecha.Value)
 Año = Year(Me.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  Me.TxtPeriodo.Text = DtaConsulta.Recordset.Periodo
  NumeroPeriodo = DtaConsulta.Recordset.NPeriodo
  NumeroTransaccion = DtaConsulta.Recordset.NTransacciones
  EstadoPeriodo = DtaConsulta.Recordset.EstadoPeriodo
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
Fecha = Format(DtaTasas.Recordset.FechaTasas, "dd/mm/yyyy")
   
    Encontrado = True
    Cambio = DtaTasas.Recordset.MontoCordobas
    MDIPrimero.StatusBar2.Panels(2) = "Tasa Dolar: " & Format(Cambio, "##,##0.00") & "    " & "Tasa Libras: " & Format(DtaTasas.Recordset.MontoLibras, "##,##0.00")
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

