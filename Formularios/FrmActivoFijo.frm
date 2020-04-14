VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmActivoFijo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Activo Fijo"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   11295
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   9960
      TabIndex        =   5
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "Borrar"
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton CmdSiguiente 
      Caption         =   "Siguiente >"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "< Anterior"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "Nuevo"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Grabar"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc DtaCuentas 
      Height          =   375
      Left            =   4320
      Top             =   8040
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
   Begin MSAdodcLib.Adodc DtaEncargado 
      Height          =   375
      Left            =   4320
      Top             =   7800
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
      Caption         =   "DtaEncargado"
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
   Begin MSAdodcLib.Adodc DtaActivoFijo 
      Height          =   375
      Left            =   480
      Top             =   8160
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
      Caption         =   "DtaActivoFijo"
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
   Begin MSAdodcLib.Adodc DtaGrupoCuentas 
      Height          =   330
      Left            =   4320
      Top             =   8160
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
      Caption         =   "DtaGrupoCuentas"
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
      Height          =   330
      Left            =   480
      Top             =   8040
      Width           =   2415
      _ExtentX        =   4260
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
   Begin MSAdodcLib.Adodc DtaBusca 
      Height          =   330
      Left            =   480
      Top             =   7800
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "DtaBusca"
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
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   5415
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11295
      _Version        =   786432
      _ExtentX        =   19923
      _ExtentY        =   9551
      _StockProps     =   68
      Appearance      =   9
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.DisableLunaColors=   0   'False
      PaintManager.OneNoteColors=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   4
      Item(0).Caption =   "Datos Generales"
      Item(0).ControlCount=   47
      Item(0).Control(0)=   "TxtDescripcion"
      Item(0).Control(1)=   "TxtLocalizacion"
      Item(0).Control(2)=   "TxtSerie"
      Item(0).Control(3)=   "TxtMarca"
      Item(0).Control(4)=   "TxtMarbete"
      Item(0).Control(5)=   "Command1"
      Item(0).Control(6)=   "DBEncargado"
      Item(0).Control(7)=   "DBCodigo"
      Item(0).Control(8)=   "TxtFechaCompra"
      Item(0).Control(9)=   "Shape1"
      Item(0).Control(10)=   "Label1"
      Item(0).Control(11)=   "Label2"
      Item(0).Control(12)=   "Label3"
      Item(0).Control(13)=   "Label5"
      Item(0).Control(14)=   "Label4"
      Item(0).Control(15)=   "Label6"
      Item(0).Control(16)=   "Label7"
      Item(0).Control(17)=   "Label8"
      Item(0).Control(18)=   "Label9"
      Item(0).Control(19)=   "TxtValorOriginal"
      Item(0).Control(20)=   "TxtValorEstMeses"
      Item(0).Control(21)=   "TxtValorRescate"
      Item(0).Control(22)=   "TxtDepAcumulada"
      Item(0).Control(23)=   "DBGrupos"
      Item(0).Control(24)=   "TxtFechaBaja"
      Item(0).Control(25)=   "TxtFechaUltDep"
      Item(0).Control(26)=   "Shape2"
      Item(0).Control(27)=   "Label10"
      Item(0).Control(28)=   "Label11"
      Item(0).Control(29)=   "Label12"
      Item(0).Control(30)=   "Label13"
      Item(0).Control(31)=   "Label14"
      Item(0).Control(32)=   "Label15"
      Item(0).Control(33)=   "Label16"
      Item(0).Control(34)=   "Command2"
      Item(0).Control(35)=   "Command3"
      Item(0).Control(36)=   "Command4"
      Item(0).Control(37)=   "Shape4"
      Item(0).Control(38)=   "Command5"
      Item(0).Control(39)=   "Label17"
      Item(0).Control(40)=   "Label18"
      Item(0).Control(41)=   "Label19"
      Item(0).Control(42)=   "Label20"
      Item(0).Control(43)=   "Label21"
      Item(0).Control(44)=   "TxtDepreciacion"
      Item(0).Control(45)=   "TxtGastos"
      Item(0).Control(46)=   "TxtCuentaOriginal"
      Item(1).Caption =   "Depreciacion"
      Item(1).ControlCount=   0
      Item(2).Caption =   "Traslados y Asignaciones"
      Item(2).ControlCount=   0
      Item(3).Caption =   "Soporte Tecnico"
      Item(3).ControlCount=   0
      Begin VB.CommandButton Command5 
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
         Left            =   10320
         Picture         =   "FrmActivoFijo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   615
         Left            =   1800
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox TxtLocalizacion 
         Height          =   285
         Left            =   1800
         TabIndex        =   21
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox TxtSerie 
         Height          =   285
         Left            =   1800
         TabIndex        =   20
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox TxtMarca 
         Height          =   495
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   3840
         Width           =   2175
      End
      Begin VB.TextBox TxtMarbete 
         Height          =   285
         Left            =   1800
         TabIndex        =   18
         Top             =   2760
         Width           =   2175
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
         Left            =   3600
         Picture         =   "FrmActivoFijo.frx":014E
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox TxtValorOriginal 
         Height          =   285
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox TxtValorEstMeses 
         Height          =   285
         Left            =   6000
         TabIndex        =   15
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox TxtValorRescate 
         Height          =   285
         Left            =   6000
         TabIndex        =   14
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox TxtDepAcumulada 
         Height          =   285
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox TxtGastos 
         Height          =   285
         Left            =   8400
         TabIndex        =   12
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox TxtCuentaOriginal 
         Height          =   285
         Left            =   8400
         TabIndex        =   11
         Top             =   2760
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
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
         Left            =   10320
         Picture         =   "FrmActivoFijo.frx":029C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1560
         Width           =   375
      End
      Begin VB.CommandButton Command3 
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
         Left            =   10320
         Picture         =   "FrmActivoFijo.frx":03EA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command4 
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
         Left            =   2040
         Picture         =   "FrmActivoFijo.frx":0538
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox TxtDepreciacion 
         Height          =   285
         Left            =   8400
         TabIndex        =   7
         Top             =   1560
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo DBEncargado 
         Height          =   315
         Left            =   1800
         TabIndex        =   23
         Top             =   4440
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DBCodigo 
         Height          =   315
         Left            =   1680
         TabIndex        =   24
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker TxtFechaCompra 
         Height          =   285
         Left            =   1800
         TabIndex        =   25
         Top             =   3120
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         _Version        =   393216
         Format          =   63832065
         CurrentDate     =   37992
      End
      Begin MSDataListLib.DataCombo DBGrupos 
         Height          =   315
         Left            =   6000
         TabIndex        =   26
         Top             =   3960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker TxtFechaBaja 
         Height          =   285
         Left            =   6000
         TabIndex        =   27
         Top             =   3000
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Format          =   63832065
         CurrentDate     =   37992
      End
      Begin MSComCtl2.DTPicker TxtFechaUltDep 
         Height          =   285
         Left            =   6000
         TabIndex        =   28
         Top             =   1560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Format          =   63832065
         CurrentDate     =   37992
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Datos Depreciacion"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   50
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Valor Original"
         Height          =   255
         Left            =   8400
         TabIndex        =   49
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta de Gastos"
         Height          =   255
         Left            =   8400
         TabIndex        =   48
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Depreciacion"
         Height          =   255
         Left            =   8400
         TabIndex        =   47
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo de Cuentas"
         Height          =   375
         Left            =   4680
         TabIndex        =   46
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BorderColor     =   &H00000000&
         DrawMode        =   2  'Blackness
         FillColor       =   &H000000FF&
         Height          =   3855
         Left            =   120
         Shape           =   5  'Rounded Square
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Datos Generales del Activo"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   44
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo Activo"
         Height          =   255
         Left            =   480
         TabIndex        =   43
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   480
         TabIndex        =   42
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Localizacion:"
         Height          =   255
         Left            =   480
         TabIndex        =   41
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Numero Marbete:"
         Height          =   255
         Left            =   480
         TabIndex        =   40
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Compra:"
         Height          =   255
         Left            =   480
         TabIndex        =   39
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Numero Serie:"
         Height          =   255
         Left            =   480
         TabIndex        =   38
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
         Height          =   255
         Left            =   480
         TabIndex        =   37
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Empleado Encag:"
         Height          =   255
         Left            =   480
         TabIndex        =   36
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         DrawMode        =   6  'Mask Pen Not
         Height          =   3855
         Left            =   4440
         Shape           =   4  'Rounded Rectangle
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Original"
         Height          =   255
         Left            =   4680
         TabIndex        =   35
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Ultima Dep"
         Height          =   255
         Left            =   4560
         TabIndex        =   34
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Estimado en Meses"
         Height          =   375
         Left            =   4680
         TabIndex        =   33
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Rescate"
         Height          =   255
         Left            =   4680
         TabIndex        =   32
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Baja"
         Height          =   255
         Left            =   4680
         TabIndex        =   31
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Depreciacion Acumulada"
         Height          =   375
         Left            =   4680
         TabIndex        =   30
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Datos Depreciacion"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   29
         Top             =   720
         Width           =   3135
      End
      Begin VB.Shape Shape4 
         DrawMode        =   6  'Mask Pen Not
         Height          =   3855
         Left            =   8280
         Shape           =   4  'Rounded Rectangle
         Top             =   1080
         Width           =   2775
      End
   End
End
Attribute VB_Name = "FrmActivoFijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAnterior_Click()
On Error GoTo TipoErrs
Dim Respuesta As Integer
DtaActivoFijo.Recordset.MovePrevious
If DtaActivoFijo.Recordset.BOF Then
   DtaActivoFijo.Recordset.MoveNext
   MsgBox "Este es el Primer Registro", vbInfoContabilidadtion, "Control de Cuentas Contabilidad"
Else
  Me.DBCodigo.Text = Me.DtaActivoFijo.Recordset!CodCuenta
End If
   Exit Sub
TipoErrs:
   ControlErrores
End Sub

Private Sub CmdBorrar_Click()
On Error GoTo TipoErrs
  Dim Respuesta, Rsp
  If DtaActivoFijo.Recordset.RecordCount = 0 Then
    MsgBox "No Existen Registros de Activos Fijos Actualmente", vbInformation
    Exit Sub
  End If
  Set Rsp = DtaActivoFijo.Recordset
  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando: " & Me.TxtDescripcion.Text)
   If Respuesta = 6 Then
     Criterio = "CodCuenta='" & Me.DBCodigo & "'"
     DtaActivoFijo.Recordset.MoveFirst
     Me.DtaActivoFijo.Recordset.Find (Criterio)
    If Not DtaActivoFijo.Recordset.EOF Then
     DtaActivoFijo.Recordset.Delete
     DtaActivoFijo.Refresh
   '/////////Borra registro de cuentas/////////////
     Criterio = "CodCuentas='" & Me.DBCodigo & "'"
     If DtaCuentas.Recordset.RecordCount <> 0 Then DtaCuentas.Recordset.MoveFirst
     Me.DtaCuentas.Recordset.Find (Criterio)
    If Not DtaCuentas.Recordset.EOF Then
      Me.DtaCuentas.Recordset.Delete
      DtaCuentas.Refresh
    End If
    End If
      Me.DBCodigo.Text = ""
  End If
  Me.DtaActivoFijo.Refresh

 Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub CmdBuscaCuenta_Click()
QueProducto = "CuentaDepreciacion"
FrmConsulta.Show 1
End Sub

Private Sub CmdBuscaGastos_Click()
QueProducto = "CuentaGastos"
FrmConsulta.Show 1
End Sub

Private Sub CmdBuscaOriginal_Click()
QueProducto = "CuentaOriginal"
FrmConsulta.Show 1
End Sub

Private Sub CmdGrabar_Click()
On Error GoTo TipoErrs

If Me.TxtCuentaOriginal.Text = "" Then
  MsgBox "Se necesita la cuenta del Activo", vbCritical, "DFID"
 Exit Sub
End If

If TxtGastos.Text = "" Then
 MsgBox "Se necesita la cuenta de Gastos", vbCritical, "DFID"
 Exit Sub
End If

If TxtDepreciacion.Text = "" Then
 MsgBox "Se necesita la cuenta para la Depreciacion", vbCritical, "DFID"
 Exit Sub
End If

If CodGrupo = "" Then
 MsgBox "Se necesita grupo de Cuenta", vbCritical, "DFID"
 Exit Sub
End If

If Me.TxtDescripcion.Text = "" Then
 MsgBox "Debe Llenar el campo de la Descripcion", vbCritical, "Sistema Contable"
 Exit Sub
End If


Criterio = "CodCuenta='" & Me.DBCodigo.Text & "'"
If DtaActivoFijo.Recordset.RecordCount <> 0 Then DtaActivoFijo.Recordset.MoveFirst
Me.DtaActivoFijo.Recordset.Find (Criterio)
If DtaActivoFijo.Recordset.EOF Then
  Criterio = "CodCuentas='" & Me.DBEncargado.Text & "'"
  If DtaCuentas.Recordset.RecordCount <> 0 Then DtaCuentas.Recordset.MoveFirst
  Me.DtaCuentas.Recordset.Find (Criterio)
  If Me.DtaCuentas.Recordset.EOF Then
'   Me.DtaCuentas.Recordset.AddNew
'   Me.DtaCuentas.Recordset("CodCuentas") = Me.DBCodigo.Text
'   Me.DtaCuentas.Recordset("DescripcionCuentas") = Me.TxtDescripcion.Text
'   Me.DtaCuentas.Recordset("TipoCuenta") = "Activo Fijo"
'   Me.DtaCuentas.Recordset.CodGrupo = CodGrupo
'    Me.DtaCuentas.Recordset("TipoMoneda") = "Dólares"
'   Me.DtaCuentas.Recordset.SaldoActual = 0#
'   Me.DtaCuentas.Recordset.Update
  End If
  Me.DtaActivoFijo.Refresh
  DtaActivoFijo.Recordset.AddNew
   DtaActivoFijo.Recordset!CodCuenta = Me.TxtCuentaOriginal.Text
   'Me.DBCodigo.Text
   DtaActivoFijo.Recordset!CuentaValorOriginal = Me.TxtCuentaOriginal.Text
   DtaActivoFijo.Recordset!CuentaGastos = Me.TxtGastos
   DtaActivoFijo.Recordset!CuentaDepreciacion = Me.TxtDepreciacion
   Me.DtaActivoFijo.Recordset!DescripcionActivo = Me.TxtDescripcion.Text
   Me.DtaActivoFijo.Recordset!Localizacion = Me.TxtLocalizacion.Text
   Me.DtaActivoFijo.Recordset!NumeroMarbete = Me.TxtMarbete.Text
   Me.DtaActivoFijo.Recordset!FechaCompra = Me.TxtFechaCompra.Value
   Me.DtaActivoFijo.Recordset!NumeroSerie = Me.TxtSerie.Text
   Me.DtaActivoFijo.Recordset!Marca = Me.TxtMarca.Text
   Me.DtaActivoFijo.Recordset!ValorOriginal = Val(Me.TxtValorOriginal.Text)
   Me.DtaActivoFijo.Recordset!FechaUltimaDepre = Me.TxtFechaUltDep.Value
   Me.DtaActivoFijo.Recordset!ValorEstimadoMeses = Val(Me.TxtValorEstMeses.Text)
   Me.DtaActivoFijo.Recordset!ValorRescate = Val(Me.TxtValorRescate.Text)
   Me.DtaActivoFijo.Recordset!FechaBaja = Me.TxtFechaBaja.Value
   'Me.DtaActivoFijo.Recordset.DepreciacionAcumulada = Me.TxtDepAcumulada.Text
 
   '/////Busco el Dodigo del Encargado/////////////////
  Criterio = "NombreEncargado='" & Me.DBEncargado.Text & "'"
  If DtaEncargado.Recordset.RecordCount <> 0 Then DtaEncargado.Recordset.MoveFirst
  Me.DtaEncargado.Recordset.Find (Criterio)
  If Not DtaEncargado.Recordset.EOF Then
    CodEncargado = DtaEncargado.Recordset!CodEncargado
  Else
    MsgBox "Seleccione de la Lista de Encargado", vbCritical, "Sistema Contable"
    Exit Sub
  End If
 
  Me.DtaActivoFijo.Recordset!CodEncargado = CodEncargado
 DtaActivoFijo.Recordset.Update
Else
  'DtaActivoFijo.Recordset.Edit
   DtaActivoFijo.Recordset!CodCuenta = Me.DBCodigo.Text
   DtaActivoFijo.Recordset!CuentaValorOriginal = Me.TxtCuentaOriginal.Text
   DtaActivoFijo.Recordset!CuentaGastos = Me.TxtGastos
   DtaActivoFijo.Recordset!CuentaDepreciacion = Me.TxtDepreciacion
   Me.DtaActivoFijo.Recordset!DescripcionActivo = Me.TxtDescripcion.Text
   Me.DtaActivoFijo.Recordset!Localizacion = Me.TxtLocalizacion.Text
   Me.DtaActivoFijo.Recordset!NumeroMarbete = Me.TxtMarbete.Text
   Me.DtaActivoFijo.Recordset!FechaCompra = Me.TxtFechaCompra.Value
   Me.DtaActivoFijo.Recordset!NumeroSerie = Me.TxtSerie.Text
   Me.DtaActivoFijo.Recordset!Marca = Me.TxtMarca.Text
   Me.DtaActivoFijo.Recordset!ValorOriginal = Me.TxtValorOriginal.Text
   Me.DtaActivoFijo.Recordset!FechaUltimaDepre = Me.TxtFechaUltDep.Value
   Me.DtaActivoFijo.Recordset!ValorEstimadoMeses = Me.TxtValorEstMeses.Text
   Me.DtaActivoFijo.Recordset!ValorRescate = Me.TxtValorRescate.Text
   Me.DtaActivoFijo.Recordset!FechaBaja = Me.TxtFechaBaja.Value
   Me.DtaActivoFijo.Recordset!DepreciacionAcumulada = Me.TxtDepAcumulada.Text
   '/////Busco el Dodigo del Encargado/////////////////
  Criterio = "NombreEncargado='" & Me.DBEncargado.Text & "'"
  Me.DtaEncargado.Recordset.Find (Criterio)
  If Not DtaEncargado.Recordset.EOF Then
    CodEncargado = DtaEncargado.Recordset!CodEncargado
  End If
 
  Me.DtaActivoFijo.Recordset!CodEncargado = CodEncargado
 
 DtaActivoFijo.Recordset.Update
End If
Me.DBCodigo.Text = ""
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub CmdNuevo_Click()
Me.DBCodigo.Text = ""
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSiguiente_Click()
On Error GoTo TipoErrs
Dim Respuesta As Integer
DtaActivoFijo.Recordset.MoveNext
If DtaActivoFijo.Recordset.EOF Then
   DtaActivoFijo.Recordset.MovePrevious
   MsgBox "Este es el Ultimo Registro", vbInfoContabilidadtion, "Control de Cuentas Contabilidad"
Else
  Me.DBCodigo.Text = Me.DtaActivoFijo.Recordset!CodCuenta
End If
   Exit Sub
TipoErrs:
   ControlErrores
End Sub

Private Sub Command1_Click()
QueProducto = "CuentaActivoFijo"
FrmConsulta.Show 1
End Sub

Private Sub DBCodigo_Change()
On Error GoTo TipoErrs
Dim Debito As Double, Credito As Double

Criterio = "CodCuenta='" & Me.DBCodigo.Text & "'"
If DtaActivoFijo.Recordset.RecordCount <> 0 Then DtaActivoFijo.Recordset.MoveFirst
Me.DtaActivoFijo.Recordset.Find (Criterio)
If DtaActivoFijo.Recordset.EOF Then
  Me.TxtDescripcion.Text = ""
  Me.TxtLocalizacion.Text = ""
  Me.TxtMarbete.Text = ""
  Me.TxtFechaCompra.Value = Format(Now, "dd/mm/yyyy")
  Me.TxtSerie.Text = ""
  Me.TxtMarca.Text = ""
  Me.TxtValorOriginal.Text = ""
  Me.TxtFechaUltDep.Value = Format(Now, "dd/mm/yyyy")
  Me.TxtValorEstMeses.Text = ""
  Me.TxtValorRescate.Text = ""
  Me.TxtFechaBaja.Value = Format(Now, "dd/mm/yyyy")
  Me.TxtDepAcumulada.Text = ""
  Me.DBEncargado.Text = ""
  Me.DBGrupos.Text = ""
  Me.TxtDepreciacion.Text = ""
  Me.TxtCuentaOriginal.Text = ""
  Me.TxtGastos.Text = ""
Else
  Me.TxtCuentaOriginal.Text = DtaActivoFijo.Recordset!CodCuenta
  Me.TxtGastos = Me.DtaActivoFijo.Recordset!CuentaGastos
  Me.TxtDepreciacion = Me.DtaActivoFijo.Recordset!CuentaDepreciacion
  Me.TxtDescripcion.Text = Me.DtaActivoFijo.Recordset!DescripcionActivo
  Me.TxtLocalizacion.Text = Me.DtaActivoFijo.Recordset!Localizacion
  Me.TxtMarbete.Text = Me.DtaActivoFijo.Recordset!NumeroMarbete
  Me.TxtFechaCompra.Value = Me.DtaActivoFijo.Recordset!FechaCompra
  Me.TxtSerie.Text = Me.DtaActivoFijo.Recordset!NumeroSerie
  Me.TxtMarca.Text = Me.DtaActivoFijo.Recordset!Marca
  'Me.TxtValorOriginal.Text = Me.DtaActivoFijo.Recordset!ValorOriginal
  Me.TxtFechaUltDep.Value = Me.DtaActivoFijo.Recordset!FechaUltimaDepre
  Me.TxtValorEstMeses.Text = Me.DtaActivoFijo.Recordset!ValorEstimadoMeses
  Me.TxtValorRescate.Text = Me.DtaActivoFijo.Recordset!ValorRescate
  Me.TxtFechaBaja.Value = Me.DtaActivoFijo.Recordset!FechaBaja
  'Me.TxtDepAcumulada.Text = Me.DtaActivoFijo.Recordset.DepreciacionAcumulada
  
'//////////////////////////////Busco el Saldo de la Depreciacion////////////////////////
 Me.DtaBusca.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Transacciones.Debito*Transacciones.TCambio) AS MDebito, Sum(Transacciones.TCambio*Transacciones.Credito) AS MCredito From Transacciones GROUP BY Transacciones.CodCuentas Having (((Transacciones.CodCuentas) = '" & Me.TxtDepreciacion.Text & "'))"
 Me.DtaBusca.Refresh
 If Not DtaBusca.Recordset.EOF Then
   Debito = Me.DtaBusca.Recordset!MDebito
   Credito = Me.DtaBusca.Recordset!MCredito
   Me.TxtDepAcumulada.Text = Credito - Debito
 Else
   Me.TxtDepAcumulada.Text = "0.00"
 End If
 
'///////////////////////Busco el Saldo del Valor Original///////////////////////////////////
 Me.DtaBusca.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Transacciones.Debito*Transacciones.TCambio) AS MDebito, Sum(Transacciones.TCambio*Transacciones.Credito) AS MCredito From Transacciones GROUP BY Transacciones.CodCuentas Having (((Transacciones.CodCuentas) = '" & Me.TxtCuentaOriginal.Text & "'))"
 Me.DtaBusca.Refresh
 If Not DtaBusca.Recordset.EOF Then
   Debito = Me.DtaBusca.Recordset!MDebito
   Credito = Me.DtaBusca.Recordset!MCredito
   Me.TxtValorOriginal.Text = Debito - Credito
 Else
   Me.TxtValorOriginal.Text = "0.00"
 End If
 

   
  '/////Busco el Dodigo del Encargado/////////////////
  CodEncargado = Me.DtaActivoFijo.Recordset!CodEncargado
  Criterio = "CodEncargado='" & CodEncargado & "'"
  If DtaEncargado.Recordset.RecordCount <> 0 Then DtaEncargado.Recordset.MoveFirst
  Me.DtaEncargado.Recordset.Find (Criterio)
  If Not DtaEncargado.Recordset.EOF Then
    Me.DBEncargado.Text = Me.DtaEncargado.Recordset!NombreEncargado
  End If
 
 '/////Busco la Descripcion del Grupo/////////////////
  CodigoCuenta = TxtCuentaOriginal
  Criterio = "CodCuentas='" & CodigoCuenta & "'"
  If DtaCuentas.Recordset.RecordCount <> 0 Then DtaCuentas.Recordset.MoveFirst
  Me.DtaCuentas.Recordset.Find (Criterio)
  If Not DtaCuentas.Recordset.EOF Then
   If Not IsNull(DtaCuentas.Recordset!CodGrupo) Then
    CodGrupo = DtaCuentas.Recordset!CodGrupo
   End If
  End If
 
 
 '/////Busco la Descripcion del Grupo/////////////////
  Criterio = "CodGrupo='" & CodGrupo & "'"
  If DtaGrupoCuentas.Recordset.RecordCount <> 0 Then DtaGrupoCuentas.Recordset.MoveFirst
  Me.DtaGrupoCuentas.Recordset.Find (Criterio)
  If Not DtaGrupoCuentas.Recordset.EOF Then
    Me.DBGrupos.Text = DtaGrupoCuentas.Recordset!DescripcionGrupo
  End If
 

End If
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub DBGrupos_Click(Area As Integer)
On Error GoTo TipoErrs:
    If DBGrupos.Text = "" Then Exit Sub 'jonathan
  Criterio = "DescripcionGrupo='" & Me.DBGrupos.Text & "'"
  If DtaGrupoCuentas.Recordset.RecordCount <> 0 Then DtaGrupoCuentas.Recordset.MoveFirst
  Me.DtaGrupoCuentas.Recordset.Find (Criterio)
  CodGrupo = DtaGrupoCuentas.Recordset!CodGrupo
  
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub Form_Activate()
Me.DtaActivoFijo.Refresh
Me.DtaCuentas.Refresh
Me.DtaEncargado.Refresh
If Not CodigoUsuario = 0 Then
 Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Activo Fijo'))"
 Me.DtaNacceso.Refresh
 If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdGrabar.Enabled = False
 End If
 Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Activo Fijo'))"
 Me.DtaNacceso.Refresh
 If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdBorrar.Enabled = False
 End If
End If


End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd
'SkinFramework.LoadSkin App.Path + "\Office2007.cjstyles", "NormalBlue.ini"
'SkinFramework.ApplyWindow Me.hWnd
'SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics

frmPane.wndTaskPanel.VisualTheme = xtpTaskPanelThemeNativeWinXP
    
Me.Shape1.BackColor = RGB(219, 226, 242)

'Me.Picture = Nothing
'Me.CmdAnterior.BackColor = RGB(208, 199, 182)
'Me.CmdBorrar.BackColor = RGB(208, 199, 182)
'Me.CmdGrabar.BackColor = RGB(208, 199, 182)
'Me.CmdNuevo.BackColor = RGB(208, 199, 182)
'Me.CmdSiguiente.BackColor = RGB(208, 199, 182)
'Me.CmdSalir.BackColor = RGB(208, 199, 182)

With Me.DtaNacceso
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from Accesos"
   .Refresh
End With

With Me.DtaGrupoCuentas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from GrupoCuentas"
   .Refresh
End With
LlenarDataCombos DtaGrupoCuentas, DBGrupos, "DescripcionGrupo", "CodGrupo"

With Me.DtaCuentas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from Cuentas"
   .Refresh
End With

With Me.DtaBusca
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaActivoFijo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from ActivoFijo"
   .Refresh
End With
LlenarDataCombos DtaActivoFijo, DBCodigo, "CodCuenta", "CodCuenta"
With Me.DtaEncargado
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from Encargado"
   .Refresh
End With
LlenarDataCombos DtaEncargado, DBEncargado, "NombreEncargado", "CodEncargado"

End Sub

Private Sub TxtCuentaOriginal_Change()
Dim Debito As Double, Credito As Double
 Me.DtaBusca.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Transacciones.Debito*Transacciones.TCambio) AS MDebito, Sum(Transacciones.TCambio*Transacciones.Credito) AS MCredito From Transacciones GROUP BY Transacciones.CodCuentas Having (((Transacciones.CodCuentas) = '" & Me.TxtCuentaOriginal.Text & "'))"
 Me.DtaBusca.Refresh
 If Not DtaBusca.Recordset.EOF Then
   Debito = Me.DtaBusca.Recordset!MDebito
   Credito = Me.DtaBusca.Recordset!MCredito
   Me.TxtValorOriginal.Text = Debito - Credito
 Else
   Me.TxtValorOriginal.Text = "0.00"
 End If
End Sub

Private Sub TxtDepreciacion_Change()
Dim Debito As Double, Credito As Double
 Me.DtaBusca.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Transacciones.Debito*Transacciones.TCambio) AS MDebito, Sum(Transacciones.TCambio*Transacciones.Credito) AS MCredito From Transacciones GROUP BY Transacciones.CodCuentas Having (((Transacciones.CodCuentas) = '" & Me.TxtDepreciacion.Text & "'))"
 Me.DtaBusca.Refresh
 If Not DtaBusca.Recordset.EOF Then
   Debito = Me.DtaBusca.Recordset!MDebito
   Credito = Me.DtaBusca.Recordset!MCredito
   Me.TxtDepAcumulada.Text = Credito - Debito
 Else
   Me.TxtDepAcumulada.Text = "0.00"
 End If
End Sub
