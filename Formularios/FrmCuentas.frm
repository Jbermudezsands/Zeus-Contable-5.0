VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmCuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   10965
   Begin MSAdodcLib.Adodc DtaCuentasCombo 
      Height          =   375
      Left            =   3480
      Top             =   7320
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
      Caption         =   "DtaCuentasCombo"
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
   Begin VB.TextBox TxtCodCuentas 
      Height          =   375
      Left            =   6480
      TabIndex        =   54
      Text            =   "Text9"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9720
      TabIndex        =   26
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton CmdSiguiente 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "Anterior"
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc DtaSaldos 
      Height          =   375
      Left            =   240
      Top             =   7080
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
      Caption         =   "DtaSaldos"
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
      Left            =   240
      Top             =   6720
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
   Begin MSAdodcLib.Adodc DtaCuentas 
      Height          =   375
      Left            =   240
      Top             =   5760
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
   Begin MSAdodcLib.Adodc DtaGrupoCuentas 
      Height          =   375
      Left            =   240
      Top             =   6360
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
   Begin MSAdodcLib.Adodc DtaConsulta 
      Height          =   375
      Left            =   240
      Top             =   6000
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   7646
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Informacion Cuentas"
      TabPicture(0)   =   "FrmCuentas.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Informacion de Saldos"
      TabPicture(1)   =   "FrmCuentas.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Informacion Cuentas"
      TabPicture(2)   =   "FrmCuentas.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).Control(2)=   "Frame6"
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame6 
         Height          =   975
         Left            =   -74880
         TabIndex        =   49
         Top             =   2640
         Width           =   5175
         Begin XtremeSuiteControls.RadioButton OptIVA 
            Height          =   255
            Left            =   2040
            TabIndex        =   53
            Top             =   240
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Cuenta IVA"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton OptRetencion 
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Cuenta Rentencion"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.TextBox TxtRetencion 
            Height          =   285
            Left            =   1080
            MaxLength       =   255
            TabIndex        =   50
            Top             =   600
            Visible         =   0   'False
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel CaptionRetencion 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmCuentas.frx":0054
            TabIndex        =   51
            Top             =   600
            Visible         =   0   'False
            Width           =   975
         End
         Begin XtremeSuiteControls.RadioButton OptNoImpuesto 
            Height          =   255
            Left            =   3360
            TabIndex        =   58
            Top             =   240
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "No Impuesto"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Datos del Representante"
         Height          =   1575
         Left            =   -70920
         TabIndex        =   43
         Top             =   480
         Width           =   6135
         Begin VB.TextBox TxtCedula 
            Height          =   285
            Left            =   1080
            MaxLength       =   255
            TabIndex        =   39
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox TxtApellido1 
            Height          =   285
            Left            =   1080
            MaxLength       =   255
            TabIndex        =   37
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox TxtApellido2 
            Height          =   285
            Left            =   3960
            MaxLength       =   255
            TabIndex        =   38
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox TxtNombre1 
            Height          =   285
            Left            =   1080
            MaxLength       =   255
            TabIndex        =   35
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox TxtNombre2 
            Height          =   285
            Left            =   3960
            MaxLength       =   255
            TabIndex        =   36
            Top             =   360
            Width           =   1935
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmCuentas.frx":00C8
            TabIndex        =   44
            Top             =   360
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
            Height          =   255
            Left            =   3240
            OleObjectBlob   =   "FrmCuentas.frx":0134
            TabIndex        =   45
            Top             =   360
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmCuentas.frx":01A0
            TabIndex        =   46
            Top             =   1080
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmCuentas.frx":0210
            TabIndex        =   47
            Top             =   720
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
            Height          =   255
            Left            =   3240
            OleObjectBlob   =   "FrmCuentas.frx":0280
            TabIndex        =   48
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Informacion de la Compañia"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   31
         Top             =   480
         Width           =   3975
         Begin VB.TextBox TxtDireccion 
            Height          =   885
            Left            =   1080
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Top             =   1080
            Width           =   2655
         End
         Begin VB.TextBox TxtTelefono 
            Height          =   285
            Left            =   1080
            MaxLength       =   255
            TabIndex        =   33
            Top             =   720
            Width           =   2655
         End
         Begin VB.TextBox TxtRUC 
            Height          =   285
            Left            =   1080
            MaxLength       =   255
            TabIndex        =   32
            Top             =   360
            Width           =   2655
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmCuentas.frx":02F0
            TabIndex        =   40
            Top             =   360
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmCuentas.frx":035A
            TabIndex        =   41
            Top             =   720
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmCuentas.frx":03C8
            TabIndex        =   42
            Top             =   1080
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   22
         Top             =   360
         Width           =   10455
         Begin TrueOleDBGrid80.TDBGrid DBGCuentas 
            Bindings        =   "FrmCuentas.frx":0438
            Height          =   3015
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   5318
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Fecha Movimiento"
            Columns(0).DataField=   "FechaTransaccion"
            Columns(0).NumberFormat=   "Short Date"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Movimiento No"
            Columns(1).DataField=   "NumeroMovimiento"
            Columns(1).NumberFormat=   "General Number"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descripcion Movimiento"
            Columns(2).DataField=   "DescripcionMovimiento"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Tasa Cambio"
            Columns(3).DataField=   "TCambio"
            Columns(3).NumberFormat=   "Fixed"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Debito"
            Columns(4).DataField=   "MDebito"
            Columns(4).NumberFormat=   "General Number"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Credito"
            Columns(5).DataField=   "MCredito"
            Columns(5).NumberFormat=   "Standard"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   -1  'True
            Splits(0).Caption=   "Historico de Saldos"
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=131588"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=131588"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=131588"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=131588"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=131588"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=131588"
            Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
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
            _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1,.fgcolor=&HFFAEFF&"
            _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HFFA8A8&,.fgcolor=&H800080&"
            _StyleDefs(20)  =   ":id=22,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(21)  =   ":id=22,.fontname=Lucida Calligraphy"
            _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bgcolor=&HECB877&"
            _StyleDefs(23)  =   ":id=14,.fgcolor=&H800000&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(24)  =   ":id=14,.strikethrough=0,.charset=0"
            _StyleDefs(25)  =   ":id=14,.fontname=MS Sans Serif"
            _StyleDefs(26)  =   "Splits(0).FooterStyle:id=15,.parent=3,.alignment=2,.bgcolor=&HFF0000&"
            _StyleDefs(27)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(28)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(29)  =   "Splits(0).EditorStyle:id=17,.parent=7,.bgcolor=&H80000005&"
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Left            =   2760
            OleObjectBlob   =   "FrmCuentas.frx":0450
            TabIndex        =   24
            Top             =   3360
            Width           =   1815
         End
         Begin ACTIVESKINLibCtl.SkinLabel LblSaldo 
            Height          =   255
            Left            =   4920
            OleObjectBlob   =   "FrmCuentas.frx":04DE
            TabIndex        =   25
            Top             =   3360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2775
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   10455
         Begin XtremeSuiteControls.CheckBox ChkCentroCostos 
            Height          =   375
            Left            =   5880
            TabIndex        =   59
            Top             =   2280
            Visible         =   0   'False
            Width           =   2895
            _Version        =   786432
            _ExtentX        =   5106
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Activar Centro de Costos"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.TextBox TxtCuentaImporta 
            Height          =   285
            Left            =   2160
            MaxLength       =   255
            TabIndex        =   56
            Top             =   2280
            Width           =   3015
         End
         Begin TrueOleDBList80.TDBCombo DBCliente 
            Bindings        =   "FrmCuentas.frx":053C
            Height          =   315
            Left            =   1080
            TabIndex        =   55
            Top             =   240
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   556
            _LayoutType     =   0
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            _DropdownWidth  =   10583
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
            _PropDict       =   $"FrmCuentas.frx":055A
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
         Begin VB.ComboBox CmbAgrupado 
            Height          =   315
            ItemData        =   "FrmCuentas.frx":0604
            Left            =   6600
            List            =   "FrmCuentas.frx":0606
            TabIndex        =   29
            Top             =   1800
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.ComboBox CmbUbicacionResultado 
            Height          =   315
            ItemData        =   "FrmCuentas.frx":0608
            Left            =   2760
            List            =   "FrmCuentas.frx":060A
            TabIndex        =   28
            Top             =   1800
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.ComboBox CmbUbicacion 
            Height          =   315
            ItemData        =   "FrmCuentas.frx":060C
            Left            =   2760
            List            =   "FrmCuentas.frx":060E
            TabIndex        =   4
            Top             =   1800
            Width           =   2415
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmCuentas.frx":0610
            TabIndex        =   27
            Top             =   1800
            Width           =   2535
         End
         Begin VB.CommandButton CmdBuscaCuenta 
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
            Left            =   6240
            Picture         =   "FrmCuentas.frx":06A6
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   240
            Width           =   375
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
            Left            =   9840
            Picture         =   "FrmCuentas.frx":07F4
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox TxtDescripcionGrupo 
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1320
            Width           =   8535
         End
         Begin VB.ComboBox CmbTipo 
            Height          =   315
            ItemData        =   "FrmCuentas.frx":0942
            Left            =   7320
            List            =   "FrmCuentas.frx":0944
            TabIndex        =   1
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox TxtDescripcion 
            Height          =   285
            Left            =   1080
            MaxLength       =   255
            TabIndex        =   2
            Top             =   840
            Width           =   8655
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmCuentas.frx":0946
            TabIndex        =   16
            Top             =   240
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmCuentas.frx":09B0
            TabIndex        =   17
            Top             =   840
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   6840
            OleObjectBlob   =   "FrmCuentas.frx":0A24
            TabIndex        =   18
            Top             =   240
            Width           =   375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmCuentas.frx":0A8A
            TabIndex        =   19
            Top             =   1320
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinAGrupado 
            Height          =   255
            Left            =   5520
            OleObjectBlob   =   "FrmCuentas.frx":0AF2
            TabIndex        =   30
            Top             =   1800
            Visible         =   0   'False
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmCuentas.frx":0B68
            TabIndex        =   57
            Top             =   2280
            Width           =   2175
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   3360
         Visible         =   0   'False
         Width           =   10455
         Begin VB.ComboBox CmbTipoMoneda 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "FrmCuentas.frx":0BF8
            Left            =   7200
            List            =   "FrmCuentas.frx":0C05
            TabIndex        =   6
            Text            =   "Córdobas"
            Top             =   240
            Width           =   2535
         End
         Begin MSDataListLib.DataCombo DBGrupos 
            Bindings        =   "FrmCuentas.frx":0C23
            Height          =   315
            Left            =   1320
            TabIndex        =   5
            Top             =   240
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "DescripcionGrupo"
            Text            =   ""
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   6000
            OleObjectBlob   =   "FrmCuentas.frx":0C41
            TabIndex        =   20
            Top             =   240
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmCuentas.frx":0CB5
            TabIndex        =   21
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin MSAdodcLib.Adodc AdoCuentas 
      Height          =   375
      Left            =   240
      Top             =   7440
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
End
Attribute VB_Name = "FrmCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChkRetencion_Click()

End Sub

Private Sub CmbTipo_Change()
Me.CmbAgrupado.Visible = False
Me.SkinAGrupado.Visible = False
Select Case Me.CmbTipo.Text
 Case "Caja"
      Me.CmbUbicacion.Text = "Cajas"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
      
 Case "Bancos"
      Me.CmbUbicacion.Text = "Bancos"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
 Case "Cuentas x Cobrar"
      Me.CmbUbicacion.Text = "Cuentas x Cobrar"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
      SSTab1.TabVisible(2) = True
 Case "Inventario"
      Me.CmbUbicacion.Text = "Inventario"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
 Case "Activo Fijo"
      Me.CmbUbicacion.Text = "Terreno y Edificios"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
 Case "Papeleria - Utiles"
      Me.CmbUbicacion.Text = "Papeleria y Utiles de Oficina"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
 Case "Otros Activos"
      Me.CmbUbicacion.Text = "Otros Activos"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
 Case "Cuentas x Pagar"
      Me.CmbUbicacion.Text = "Proveedores"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
      SSTab1.TabVisible(2) = True
 Case "Pasivo"
      Me.CmbUbicacion.Text = "Pasivos Acumulados"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
 Case "Otros Pasivos"
      Me.CmbUbicacion.Text = "Otros Pasivos"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
 Case "Capital"
      Me.CmbUbicacion.Text = "Acciones Comunes"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
 Case "Ingresos - Ventas"
      Me.CmbUbicacionResultado.Text = "Ingresos - Ventas"
      Me.CmbUbicacion.Visible = False
      Me.CmbUbicacionResultado.Visible = True
 Case "Costos"
      Me.CmbUbicacionResultado.Text = "Costos"
      Me.CmbUbicacion.Visible = False
      Me.CmbUbicacionResultado.Visible = True
      Me.CmbAgrupado.Visible = True
      Me.SkinAGrupado.Visible = True
 Case "Gastos"
      Me.CmbUbicacionResultado.Text = "Gastos"
      Me.CmbUbicacion.Visible = False
      Me.CmbUbicacionResultado.Visible = True
      Me.CmbAgrupado.Visible = True
      Me.SkinAGrupado.Visible = True
 Case Else: Me.CmbUbicacion.Text = " "
   
 End Select
End Sub

Private Sub CmbTipo_Click()
Me.CmbAgrupado.Visible = False
Me.SkinAGrupado.Visible = False
Me.ChkCentroCostos.Visible = False
Select Case Me.CmbTipo.Text
 Case "Caja"
      Me.CmbUbicacion.Text = "Cajas"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
      
 Case "Bancos"
      Me.CmbUbicacion.Text = "Bancos"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
 Case "Cuentas x Cobrar"
      Me.CmbUbicacion.Text = "Cuentas x Cobrar"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
      SSTab1.TabVisible(2) = True
      Me.ChkCentroCostos.Visible = True
 Case "Inventario"
      Me.CmbUbicacion.Text = "Inventario"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
      Me.ChkCentroCostos.Visible = True
 Case "Activo Fijo"
      Me.CmbUbicacion.Text = "Terreno y Edificios"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
 Case "Papeleria - Utiles"
      Me.CmbUbicacion.Text = "Papeleria y Utiles de Oficina"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
 Case "Otros Activos"
      Me.CmbUbicacion.Text = "Otros Activos"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
      Me.ChkCentroCostos.Visible = True
 Case "Cuentas x Pagar"
      Me.CmbUbicacion.Text = "Proveedores"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
      SSTab1.TabVisible(2) = True
 Case "Pasivo"
      Me.CmbUbicacion.Text = "Pasivos Acumulados"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
 Case "Otros Pasivos"
      Me.CmbUbicacion.Text = "Otros Pasivos"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
 Case "Capital"
      Me.CmbUbicacion.Text = "Acciones Comunes"
      Me.CmbUbicacion.Visible = True
      Me.CmbUbicacionResultado.Visible = False
 Case "Ingresos - Ventas"
      Me.CmbUbicacionResultado.Text = "Ingresos - Ventas"
      Me.CmbUbicacion.Visible = False
      Me.CmbUbicacionResultado.Visible = True
 Case "Costos"
      Me.CmbUbicacionResultado.Text = "Costos"
      Me.CmbUbicacion.Visible = False
      Me.CmbUbicacionResultado.Visible = True
      Me.CmbAgrupado.Visible = True
      Me.SkinAGrupado.Visible = True
 Case "Gastos"
      Me.CmbUbicacionResultado.Text = "Gastos"
      Me.CmbUbicacion.Visible = False
      Me.CmbUbicacionResultado.Visible = True
      Me.CmbAgrupado.Visible = True
      Me.SkinAGrupado.Visible = True
 Case Else: Me.CmbUbicacion.Text = " "
   
 End Select
End Sub

Private Sub CmbUbicacion_Click()

Select Case Me.CmbUbicacion.Text

 Case "***ACTIVO CIRCULANTE***": Me.CmbUbicacion.Text = "Caja"
 Case "***ACTIVO FIJO***": Me.CmbUbicacion.Text = "Terreno y Edificios"
 Case "***ACTIVO DIFERIDO***": Me.CmbUbicacion.Text = "Papeleria y Utiles de Oficina"
 Case "***PASIVO CIRCULANTE***": Me.CmbUbicacion.Text = "Proveedores"
 Case "***PASIVO FIJO***": Me.CmbUbicacion.Text = "Cuentas x Pagar LP"
 Case "***PASIVO DIFERIDO***": Me.CmbUbicacion.Text = "Otros Pasivos"
 Case "***CAPITAL SOCIAL***": Me.CmbUbicacion.Text = "Acciones Comunes"
End Select
End Sub

Private Sub CmbUbicacionResultado_Click()
Select Case Me.CmbUbicacionResultado.Text

 Case "***INGRESOS Y VENTAS***": Me.CmbUbicacionResultado.Text = "Ingresos - Ventas"
 Case "***COSTOS Y GASTOS***": Me.CmbUbicacionResultado.Text = "Compras"
 
End Select
End Sub

Private Sub CmdAnterior_Click()
On Error GoTo TipoErrs
Dim Respuesta As Integer
DtaCuentas.Recordset.MovePrevious
If DtaCuentas.Recordset.BOF Then
   DtaCuentas.Recordset.MoveNext
   MsgBox "Este es el Primer Registro", vbInfoContabilidadtion, "Control de Cuentas Contabilidad"
Else
  Me.DBCliente.Text = DtaCuentas.Recordset("CodCuentas")
End If
   Exit Sub
TipoErrs:
   ControlErrores
End Sub

Private Sub CmdBorrar_Click()
On Error GoTo TipoErrs
  Dim Respuesta, Rsp
  Me.DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.CodGrupo, Cuentas.SaldoActual, Cuentas.TipoMoneda, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo From Cuentas Where (((Cuentas.CodCuentas) = '" & Me.DBCliente.Text & "'))"
  Me.DtaConsulta.Refresh
  
  If Not DtaConsulta.Recordset.EOF Then
     Set Rsp = DtaCuentas.Recordset
     TipoMoneda = Me.DtaConsulta.Recordset("TipoMoneda")
     Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando: " & Me.DBCliente.Text)
     If Respuesta = 6 Then
   Me.DtaSaldos.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento, Transacciones.DescripcionMovimiento,Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Debito*TCambio AS MDebito, TCambio*Credito AS MCredito From Transacciones Where (((Transacciones.CodCuentas) =  '" & Me.DBCliente.Text & "')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
        Me.DtaSaldos.Refresh
        If DtaSaldos.Recordset.EOF Then
         DtaConsulta.Recordset.Delete
        Else
          FrmTransferencia.Txtorigen.Text = Me.DBCliente.Text
          FrmTransferencia.Show 1
        End If
        
      Me.DBCliente.Text = ""
     End If
  End If
' Me.DtaCuentas.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas From Cuentas Where (((Cuentas.KeyGrupo) = '" & KeyPrincipal & "')) ORDER BY Cuentas.CodCuentas"
 
 Me.DtaCuentas.Refresh
'  Me.DBGCuentas.Columns(5).Visible = False
'  Me.DBGCuentas.Columns(6).Visible = False
'  Me.DBGCuentas.Columns(7).Caption = "Debito"
'  Me.DBGCuentas.Columns(8).Caption = "Credito"
'  Me.DBGCuentas.Columns(7).NumberFormat = "##,##0.00"
'  Me.DBGCuentas.Columns(8).NumberFormat = "##,##0.00"
'  Me.DBGCuentas.Columns(4).NumberFormat = "##,##0.000000"
  Me.LblSaldo.Caption = Format(Total1, "##,##0.00")
' Me.DBGCuentas.Columns(0).Visible = False


 Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub CmdBuscaCuenta_Click()
QueProducto = "Cuenta"
FrmConsulta.Show 1
End Sub

Private Sub CmdBuscarEmpleado_Click()
QUIEN = "Cuentas"
FrmGrupoLista.Show 1
End Sub

Private Sub CmdGrabar_Click()
Dim UbicacionReporte As String
On Error GoTo TipoErrs


  If Me.CmbTipoMoneda.Text = "" Then
    MsgBox "Se necesita el Tipo de Moneda", vbCritical, "Sistema Contable"
    Exit Sub
  End If
  
  If Me.TxtDescripcionGrupo.Text = "" Then
    MsgBox "Debe Seleccionar un Grupo", vbCritical, "Sistema Contable"
    Exit Sub
  End If

If KeyGrupoCuenta = "" Then
    MsgBox "Debe Seleccionar un Grupo", vbCritical, "Sistema Contable"
    Exit Sub
End If

'//////////////////////////////////////////////////////////////////////
'//////////VERIFICO SI LA UBICACION ESTA CORRECTA//////////////////////
'//////////////////////////////////////////////////////////////////////
If Me.CmbUbicacion.Text = "" Then
    Select Case Me.CmbTipo.Text
     Case "Caja": Me.CmbUbicacion.Text = "Cajas"
     Case "Bancos": Me.CmbUbicacion.Text = "Bancos"
     Case "Cuentas x Cobrar": Me.CmbUbicacion.Text = "Cuentas x Cobrar"
     Case "Inventario": Me.CmbUbicacion.Text = "Inventario"
     Case "Activo Fijo": Me.CmbUbicacion.Text = "Terreno y Edificios"
     Case "Papeleria - Utiles": Me.CmbUbicacion.Text = "Papeleria y Utiles de Oficina"
     Case "Otros Activos": Me.CmbUbicacion.Text = "Otros Activos"
     Case "Cuentas x Pagar": Me.CmbUbicacion.Text = "Proveedores"
     Case "Pasivo": Me.CmbUbicacion.Text = "Pasivos Acumulados"
     Case "Otros Pasivos": Me.CmbUbicacion.Text = "Otros Pasivos"
     Case "Capital": Me.CmbUbicacion.Text = "Acciones Comunes"
     Case Else: Me.CmbUbicacion.Text = " "
    End Select
End If

If Me.CmbTipo.Text = "Ingresos - Ventas" Or Me.CmbTipo.Text = "Costos" Or Me.CmbTipo.Text = "Gastos" Then
  UbicacionReporte = Me.CmbUbicacionResultado.Text
Else
   UbicacionReporte = Me.CmbUbicacion.Text
End If



Criterio = "CodCuentas='" & Me.DBCliente.Text & "'"
Me.DtaCuentas.Recordset.Find (Criterio)
If DtaCuentas.Recordset.EOF Then
 DtaCuentas.Recordset.AddNew
  DtaCuentas.Recordset("KeyGrupo") = KeyGrupoCuenta
  DtaCuentas.Recordset!DescripcionGrupo = Me.TxtDescripcionGrupo.Text
  DtaCuentas.Recordset("CodCuentas") = Me.DBCliente.Text
  DtaCuentas.Recordset("DescripcionCuentas") = Me.TxtDescripcion
  DtaCuentas.Recordset("TipoCuenta") = Me.CmbTipo.Text
  If Me.TxtCuentaImporta.Text <> "" Then
   DtaCuentas.Recordset("CodCuentaImporta") = Me.TxtCuentaImporta.Text
  Else
   DtaCuentas.Recordset("CodCuentaImporta") = Me.DBCliente.Text
  End If
  
 If Not Me.CmbAgrupado.Text = "" Then
   DtaCuentas.Recordset("SubDivicion") = Me.CmbAgrupado.Text
  End If
  
  If Not UbicacionReporte = "" Then
   
   DtaCuentas.Recordset("UbicacionReporte") = UbicacionReporte
  End If
 
  If CodGrupo <> "" Then
   DtaCuentas.Recordset("CodGrupo") = CodGrupo
  End If
  DtaCuentas.Recordset("TipoMoneda") = CmbTipoMoneda
  DtaCuentas.Recordset("SaldoActual") = 0#
  
   If Me.OptRetencion.Value = True Then
   DtaCuentas.Recordset("CausaRetencion") = True
   DtaCuentas.Recordset("CausaIva") = False
 Else
   DtaCuentas.Recordset("CausaRetencion") = False
'   DtaCuentas.Recordset("CausaIva") = True
 End If

  If Me.OptIva.Value = True Then
   DtaCuentas.Recordset("CausaIva") = True
   DtaCuentas.Recordset("CausaRetencion") = False
  Else
   DtaCuentas.Recordset("CausaIva") = False
'   DtaCuentas.Recordset("CausaRetencion") = True
  End If
  
  If Not Me.TxtRetencion.Text = "" Then
   DtaCuentas.Recordset("DescRetencion") = Me.TxtRetencion.Text
  End If

  
  If Not Me.TxtNombre1.Text = "" Then
   DtaCuentas.Recordset("Nombre1") = Me.TxtNombre1.Text
  End If

  If Not Me.TxtNombre2.Text = "" Then
   DtaCuentas.Recordset("Nombre2") = Me.TxtNombre2.Text
  End If
  
  If Not Me.TxtApellido1.Text = "" Then
   DtaCuentas.Recordset("Apellido1") = Me.TxtApellido1.Text
  End If
  
  If Not Me.TxtApellido2.Text = "" Then
   DtaCuentas.Recordset("Apellido2") = Me.TxtApellido2.Text
  End If
  
  If Not Me.txtcedula.Text = "" Then
   DtaCuentas.Recordset("Cedula") = Me.txtcedula.Text
  End If
  
  If Not Me.TxtRUC.Text = "" Then
   DtaCuentas.Recordset("RUC") = Me.TxtRUC.Text
  End If
  
  If Not Me.TxtTelefono.Text = "" Then
   DtaCuentas.Recordset("Telefono") = Me.TxtTelefono.Text
  End If
  
  If Not Me.TxtDireccion.Text = "" Then
   DtaCuentas.Recordset("Direccion") = Me.TxtDireccion.Text
  End If
  
     If Me.ChkCentroCostos.Value = xtpChecked Then
      DtaCuentas.Recordset("CentroCostos") = "True"
     Else
      DtaCuentas.Recordset("CentroCostos") = "False"
     End If

 DtaCuentas.Recordset.Update
Else
  'DtaCuentas.Recordset.Edit
  DtaCuentas.Recordset("KeyGrupo") = KeyGrupoCuenta
  DtaCuentas.Recordset("DescripcionGrupo") = Me.TxtDescripcionGrupo.Text
  DtaCuentas.Recordset("DescripcionCuentas") = Me.TxtDescripcion
  DtaCuentas.Recordset("TipoCuenta") = Me.CmbTipo.Text
  
 If Me.TxtCuentaImporta.Text <> "" Then
   DtaCuentas.Recordset("CodCuentaImporta") = Me.TxtCuentaImporta.Text
  Else
   DtaCuentas.Recordset("CodCuentaImporta") = Me.DBCliente.Text
  End If
  
  If Not Me.CmbAgrupado.Text = "" Then
   DtaCuentas.Recordset("SubDivicion") = Me.CmbAgrupado.Text
  End If
  

 If Me.OptRetencion.Value = True Then
   DtaCuentas.Recordset("CausaRetencion") = True
   DtaCuentas.Recordset("CausaIva") = False
 Else
   DtaCuentas.Recordset("CausaRetencion") = False
'   DtaCuentas.Recordset("CausaIva") = True
 End If

  If Me.OptIva.Value = True Then
   DtaCuentas.Recordset("CausaIva") = True
   DtaCuentas.Recordset("CausaRetencion") = False
  Else
   DtaCuentas.Recordset("CausaIva") = False
'   DtaCuentas.Recordset("CausaRetencion") = True
  End If
  
  If Not Me.TxtRetencion.Text = "" Then
   DtaCuentas.Recordset("DescRetencion") = Me.TxtRetencion.Text
  End If

  
  If Not Me.TxtNombre1.Text = "" Then
   DtaCuentas.Recordset("Nombre1") = Me.TxtNombre1.Text
  End If

  If Not Me.TxtNombre2.Text = "" Then
   DtaCuentas.Recordset("Nombre2") = Me.TxtNombre2.Text
  End If
  
  If Not Me.TxtApellido1.Text = "" Then
   DtaCuentas.Recordset("Apellido1") = Me.TxtApellido1.Text
  End If
  
  If Not Me.TxtApellido2.Text = "" Then
   DtaCuentas.Recordset("Apellido2") = Me.TxtApellido2.Text
  End If
  
  If Not Me.txtcedula.Text = "" Then
   DtaCuentas.Recordset("Cedula") = Me.txtcedula.Text
  End If
  
  If Not Me.TxtRUC.Text = "" Then
   DtaCuentas.Recordset("RUC") = Me.TxtRUC.Text
  End If
  
  If Not Me.TxtTelefono.Text = "" Then
   DtaCuentas.Recordset("Telefono") = Me.TxtTelefono.Text
  End If
  
  If Not Me.TxtDireccion.Text = "" Then
   DtaCuentas.Recordset("Direccion") = Me.TxtDireccion.Text
  End If

       If Me.ChkCentroCostos.Value = xtpChecked Then
      DtaCuentas.Recordset("CentroCostos") = "True"
     Else
      DtaCuentas.Recordset("CentroCostos") = "False"
     End If
  
  If Not Me.CmbUbicacion.Text = "" Then
   DtaCuentas.Recordset("UbicacionReporte") = UbicacionReporte
  End If
  
  If CodGrupo <> "" Then
   DtaCuentas.Recordset("CodGrupo") = CodGrupo
  End If
  DtaCuentas.Recordset("TipoMoneda") = CmbTipoMoneda
 DtaCuentas.Recordset.Update
End If
DtaCuentas.Refresh
   SSTab1.TabVisible(2) = False
   Me.DBCliente.Text = ""
   Me.CmbAgrupado.Text = ""
   Me.OptRetencion.Value = False
   Me.OptIva.Value = False
   Me.TxtRetencion.Text = ""
   Me.TxtNombre1.Text = ""
   Me.TxtNombre2.Text = ""
   Me.TxtApellido1.Text = ""
   Me.TxtApellido2.Text = ""
   Me.txtcedula.Text = ""
   Me.TxtRUC.Text = ""
   Me.TxtTelefono.Text = ""
   Me.TxtDireccion.Text = ""


  Me.CmbTipo.Enabled = True
  Me.TxtDescripcionGrupo.Text = ""
  Me.TxtDescripcion.Text = ""
  Me.DBGrupos.Text = ""
  Me.LblSaldo.Caption = "0.00"

  Me.DtaSaldos.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento, Transacciones.DescripcionMovimiento, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Debito*TCambio AS MDebito, TCambio*Credito AS MCredito From Transacciones Where (((Transacciones.CodCuentas) =  '-1')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
  Me.DtaSaldos.Refresh
  Me.LblSaldo.Caption = Format(Total1, "##,##0.00")

  Exit Sub
Exit Sub
TipoErrs:
'   MsgBox err.Description, vbCritical, err.Number
   ControlErrores
End Sub

Private Sub CmdNuevo_Click()
Me.DBCliente.Text = ""
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSiguiente_Click()
On Error GoTo TipoErrs
Dim Respuesta As Integer
DtaCuentas.Recordset.MoveNext
If DtaCuentas.Recordset.EOF Then
   DtaCuentas.Recordset.MovePrevious
   MsgBox "Este es el Ultimo Registro", vbInfoContabilidadtion, "Control de Cuentas Contabilidad"
Else
  Me.DBCliente.Text = DtaCuentas.Recordset("CodCuentas")
End If
   Exit Sub
TipoErrs:
   ControlErrores
End Sub


Private Sub DBCliente_Change()
 Me.TxtCodCuentas.Text = Me.DBCliente.Text
End Sub

Private Sub DBCliente_ItemChange()
 Me.TxtCodCuentas.Text = Me.DBCliente.Text
End Sub

Private Sub DBCliente_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   Me.TxtCodCuentas.Text = Me.DBCliente.Text
 End If
End Sub

Private Sub DBCliente_SelChange(Cancel As Integer)
 Me.TxtCodCuentas.Text = Me.DBCliente.Text
End Sub

Private Sub DBGrupos_Change()
On Error GoTo TipoErrs:
If Me.DBGrupos.Text = "" Then
  Exit Sub
End If
  Me.DtaGrupoCuentas.Refresh
  Criterio = "DescripcionGrupo='" & Me.DBGrupos.Text & "'"
  Me.DtaGrupoCuentas.Recordset.Find Criterio
  CodGrupo = DtaGrupoCuentas.Recordset("CodGrupo")
  
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub Form_Activate()
 On Error GoTo TipoErrs
' Me.DtaGrupoCuentas.Refresh
 Me.DtaSaldos.Refresh
 Me.DtaCuentas.Refresh
 
'  Me.DBGCuentas.Columns(5).Visible = False
'  Me.DBGCuentas.Columns(6).Visible = False
'  Me.DBGCuentas.Columns(7).Caption = "Debito"
'  Me.DBGCuentas.Columns(8).Caption = "Credito"
'  Me.DBGCuentas.Columns(7).NumberFormat = "##,##0.00"
'  Me.DBGCuentas.Columns(8).NumberFormat = "##,##0.00"
'  Me.DBGCuentas.Columns(4).NumberFormat = "##,##0.000000"
'  Me.DBGCuentas.Columns(0).Visible = False

Me.CmbAgrupado.AddItem ("Fijo")
Me.CmbAgrupado.AddItem ("Variable")


Me.CmbUbicacion.AddItem ("***ACTIVO CIRCULANTE***")
Me.CmbUbicacion.AddItem ("Cajas")
Me.CmbUbicacion.AddItem ("Bancos")
Me.CmbUbicacion.AddItem ("Cuentas x Cobrar")
Me.CmbUbicacion.AddItem ("Inventario")
Me.CmbUbicacion.AddItem ("***ACTIVO FIJO***")
Me.CmbUbicacion.AddItem ("Terreno y Edificios")
Me.CmbUbicacion.AddItem ("Mobiliario y Equipo de Oficina")
Me.CmbUbicacion.AddItem ("Equipo Rodante")
Me.CmbUbicacion.AddItem ("Depreciacion Acumulada")
Me.CmbUbicacion.AddItem ("***ACTIVO DIFERIDO***")
Me.CmbUbicacion.AddItem ("Papeleria y Utiles de Oficina")
Me.CmbUbicacion.AddItem ("Pagos Anticipados")
Me.CmbUbicacion.AddItem ("Otros Activos")
Me.CmbUbicacion.AddItem ("***PASIVO CIRCULANTE***")
Me.CmbUbicacion.AddItem ("Proveedores")
Me.CmbUbicacion.AddItem ("Impuestos x Pagar")
Me.CmbUbicacion.AddItem ("Documentos x Pagar CP")
Me.CmbUbicacion.AddItem ("Cobros Anticipados")
Me.CmbUbicacion.AddItem ("Pasivos Acumulados")
Me.CmbUbicacion.AddItem ("***PASIVO FIJO***")
Me.CmbUbicacion.AddItem ("Cuentas x Pagar LP")
Me.CmbUbicacion.AddItem ("Documentos x Pagar LP")
Me.CmbUbicacion.AddItem ("***PASIVO DIFERIDO***")
Me.CmbUbicacion.AddItem ("Otros Pasivos")
Me.CmbUbicacion.AddItem ("***CAPITAL SOCIAL***")
Me.CmbUbicacion.AddItem ("Acciones Comunes")
Me.CmbUbicacion.AddItem ("Utilidades Acumuladas")
Me.CmbUbicacion.AddItem ("Otras Ctas de Capital")

Me.CmbUbicacionResultado.AddItem ("***INGRESOS Y VENTAS***")
Me.CmbUbicacionResultado.AddItem ("Ingresos - Ventas")
Me.CmbUbicacionResultado.AddItem ("Servicios - Ventas")
Me.CmbUbicacionResultado.AddItem ("Comision - Ventas")
Me.CmbUbicacionResultado.AddItem ("Rebajas y Dev S/Venta")
Me.CmbUbicacionResultado.AddItem ("***COSTOS Y GASTOS***")
Me.CmbUbicacionResultado.AddItem ("Compras")
Me.CmbUbicacionResultado.AddItem ("Costos")
Me.CmbUbicacionResultado.AddItem ("Costos Produccion")
Me.CmbUbicacionResultado.AddItem ("Costos Generales Produccion")
Me.CmbUbicacionResultado.AddItem ("Acarreo y Fletes")
Me.CmbUbicacionResultado.AddItem ("Rebajas y Dev S/Compra")
Me.CmbUbicacionResultado.AddItem ("Sueldos y Comisiones")
Me.CmbUbicacionResultado.AddItem ("Propaganda")
Me.CmbUbicacionResultado.AddItem ("Gastos")
Me.CmbUbicacionResultado.AddItem ("Sueldos Admon")
Me.CmbUbicacionResultado.AddItem ("Energia y Agua Potable")
Me.CmbUbicacionResultado.AddItem ("Comisiones/Intereses Gandados")
Me.CmbUbicacionResultado.AddItem ("Comisiones/Intereses Pagados")
Me.CmbUbicacionResultado.AddItem ("Otros Ingresos")
Me.CmbUbicacionResultado.AddItem ("Otros Gastos")
Me.CmbUbicacionResultado.AddItem ("Impuestos Pagados")

 
 Me.CmbTipo.AddItem ("Otros Activos")
 Me.CmbTipo.AddItem ("Caja")
 Me.CmbTipo.AddItem ("Bancos")
 Me.CmbTipo.AddItem ("Cuentas x Cobrar")
 Me.CmbTipo.AddItem ("Inventario")
 Me.CmbTipo.AddItem ("Papeleria - Utiles")
 Me.CmbTipo.AddItem ("Activo Fijo")
 Me.CmbTipo.AddItem ("Otros Pasivos")
 Me.CmbTipo.AddItem ("Cuentas x Pagar")
 Me.CmbTipo.AddItem ("Pasivo")
 Me.CmbTipo.AddItem ("Capital")
 Me.CmbTipo.AddItem ("Ingresos - Ventas")
 Me.CmbTipo.AddItem ("Costos")
 Me.CmbTipo.AddItem ("Gastos")
 Me.CmbTipo.AddItem ("Cuentas de Orden")
 
If Not CodigoUsuario = 0 Then
 Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Cuentas'))"
 Me.DtaNacceso.Refresh
 If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdGrabar.Enabled = False
 End If
 Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Cuentas'))"
 Me.DtaNacceso.Refresh
 If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdBorrar.Enabled = False
 End If
End If
Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd
 Me.DBGCuentas.EvenRowStyle.BackColor = &H80FFFF
 Me.DBGCuentas.OddRowStyle.BackColor = &HC0FFFF
 Me.DBGCuentas.AlternatingRowStyle = True


SSTab1.TabVisible(2) = False

With Me.DtaConsulta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
'   .RecordSource = "Cuentas"
'   .Refresh
End With


With Me.DtaNacceso
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
'   .RecordSource = "Cuentas"
'   .Refresh
End With

With Me.DtaCuentas
   .ConnectionString = Conexion
   .RecordSource = "Cuentas"
   .Refresh
End With

With Me.DtaCuentasCombo
   .ConnectionString = Conexion
End With

With Me.DtaGrupoCuentas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "GrupoCuentas"
   .Refresh
End With

With Me.DtaSaldos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
'   .RecordSource = "Cuentas"
'   .Refresh
End With

Me.DtaCuentasCombo.RecordSource = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta From Cuentas"
Me.DtaCuentasCombo.Refresh



CambiarTipo = True
'Me.DBCliente.ListField = "CodCuentas"


Me.DtaSaldos.RecordSource = "SELECT CodCuentas, FechaTransaccion, NumeroMovimiento, DescripcionMovimiento, Debito, Credito, Debito * TCambio AS MDebito,TCambio * Credito AS MCredito, TCambio From Transacciones WHERE (CodCuentas = '-111') ORDER BY FechaTransaccion, NumeroMovimiento"
Me.DtaSaldos.Refresh

  Me.DBGCuentas.Columns(4).NumberFormat = "##,##0.00"
  Me.DBGCuentas.Columns(5).NumberFormat = "##,##0.00"

End Sub

Private Sub OptIVA_Click()
If Me.OptRetencion.Value = True Then
 Me.TxtRetencion.Visible = True
 Me.CaptionRetencion.Visible = True
 Me.CaptionRetencion.Caption = "%Retencion"
 Me.Frame4.Visible = False
 Me.Frame5.Visible = False
Else
 Me.TxtRetencion.Visible = True
 Me.CaptionRetencion.Visible = True
 Me.CaptionRetencion.Caption = "%I.V.A"
End If
End Sub

Private Sub OptNoImpuesto_Click()
 Me.Frame4.Visible = True
 Me.Frame5.Visible = True
 Me.CaptionRetencion.Visible = False
 Me.TxtRetencion.Visible = False
End Sub

Private Sub OptRetencion_Click()
If Me.OptRetencion.Value = True Then
 Me.TxtRetencion.Visible = True
 Me.CaptionRetencion.Visible = True
 Me.CaptionRetencion.Caption = "%Retencion"
 Me.Frame4.Visible = False
 Me.Frame5.Visible = False
Else
 Me.TxtRetencion.Visible = True
 Me.CaptionRetencion.Visible = True
 Me.CaptionRetencion.Caption = "%I.V.A"
End If
End Sub

Private Sub TDBCombo1_ItemChange()

'   Me.TxtCodCuentas.Text = Me.TDBCombo1.Text

End Sub



Private Sub TxtCodCuentas_Change()

 
 Dim Rbusqueda As Boolean
On Error GoTo TipoErrs
Dim Debito As Double, Credito As Double
Dim KeyGrupo As String

Total1 = 0
If Me.DBCliente.Text = "" Then
   Me.CmbAgrupado.Text = ""
   Me.OptRetencion.Value = False
   Me.OptIva.Value = False
   Me.TxtRetencion.Text = ""
   Me.TxtNombre1.Text = ""
   Me.TxtNombre2.Text = ""
   Me.TxtApellido1.Text = ""
   Me.TxtApellido2.Text = ""
   Me.txtcedula.Text = ""
   Me.TxtRUC.Text = ""
   Me.TxtTelefono.Text = ""
   Me.TxtDireccion.Text = ""
   Me.TxtCuentaImporta.Text = ""
  Me.CmbTipo.Enabled = True
  Me.TxtDescripcionGrupo.Text = ""
  Me.TxtDescripcion.Text = ""
  Me.DBGrupos.Text = ""
  Me.LblSaldo.Caption = "0.00"
  Me.ChkCentroCostos.Visible = False

  Me.DtaSaldos.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento, Transacciones.DescripcionMovimiento, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Debito*TCambio AS MDebito, TCambio*Credito AS MCredito From Transacciones Where (((Transacciones.CodCuentas) =  '-1')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
  Me.DtaSaldos.Refresh
  Me.LblSaldo.Caption = Format(Total1, "##,##0.00")

  Exit Sub
End If

Me.DtaCuentas.Refresh
Criterio = "CodCuentas='" & Me.DBCliente.Text & "'"
 Me.DtaCuentas.Recordset.Find (Criterio)
If Me.DtaCuentas.Recordset.EOF Then
'If kk = 1 Then
  Me.CmbTipo.Enabled = True
  Me.TxtDescripcionGrupo.Text = ""
  Me.TxtDescripcion.Text = ""
  Me.TxtCuentaImporta.Text = ""
'  Me.CmbTipo.Text = ""
  Me.DBGrupos.Text = ""
  Me.LblSaldo.Caption = "0.00"
'  Me.CmbTipoMoneda.Text = ""
  Me.DtaSaldos.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento, Transacciones.DescripcionMovimiento, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Debito*TCambio AS MDebito, TCambio*Credito AS MCredito From Transacciones Where (((Transacciones.CodCuentas) =  '-1')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
  Me.DtaSaldos.Refresh
'  Me.DBGCuentas.Columns(5).Visible = False
'  Me.DBGCuentas.Columns(6).Visible = False
'  Me.DBGCuentas.Columns(7).Caption = "Debito"
'  Me.DBGCuentas.Columns(8).Caption = "Credito"
'  Me.DBGCuentas.Columns(7).NumberFormat = "##,##0.00"
'  Me.DBGCuentas.Columns(8).NumberFormat = "##,##0.00"
'  Me.DBGCuentas.Columns(4).NumberFormat = "##,##0.000000"
  Me.LblSaldo.Caption = Format(Total1, "##,##0.00")
' Me.DBGCuentas.Columns(0).Visible = False
'  Me.CmbTipoMoneda.Enabled = True
  
Else

'Me.DtaCuentas.Refresh
 
' If Not IsNull(DtaCuentas.Recordset("DescripcionGrupo")) Then
'  Me.TxtDescripcionGrupo.Text = DtaCuentas.Recordset("DescripcionGrupo")
' End If
  
  Me.TxtDescripcion.Text = DtaCuentas.Recordset("DescripcionCuentas")
'
  If Not IsNull(DtaCuentas.Recordset("TipoCuenta")) Then
   TipoCuenta = DtaCuentas.Recordset("TipoCuenta")
   Me.CmbTipo.Text = DtaCuentas.Recordset("TipoCuenta")
  End If
  
  Select Case TipoCuenta
   Case "Cuentas x Cobrar"
      SSTab1.TabVisible(2) = True
   Case "Cuentas x Pagar"
      SSTab1.TabVisible(2) = True
  End Select
  
  If Not IsNull(DtaCuentas.Recordset("CausaIva")) Then
   If DtaCuentas.Recordset("CausaIva") = True Then
     Me.OptIva.Value = True
   Else
    Me.OptIva.Value = False
   End If
  End If
  
  If Not IsNull(DtaCuentas.Recordset("CausaRetencion")) Then
   If DtaCuentas.Recordset("CausaRetencion") = True Then
     Me.OptRetencion.Value = True
   Else
     Me.OptRetencion.Value = False
   End If
  End If
  
  If Me.OptIva.Value = False Then
   If Me.OptRetencion.Value = False Then
     Me.OptNoImpuesto.Value = True
   End If
  End If
  
  If Not IsNull(DtaCuentas.Recordset("DescRetencion")) Then
   Me.TxtRetencion.Text = DtaCuentas.Recordset("DescRetencion")
  End If
  
  If Not IsNull(DtaCuentas.Recordset("Nombre1")) Then
   Me.TxtNombre1.Text = DtaCuentas.Recordset("Nombre1")
  End If
  
  If Not IsNull(DtaCuentas.Recordset("Nombre2")) Then
   Me.TxtNombre2.Text = DtaCuentas.Recordset("Nombre2")
  End If
  
  If Not IsNull(DtaCuentas.Recordset("Apellido1")) Then
   Me.TxtApellido1.Text = DtaCuentas.Recordset("Apellido1")
  End If
  
  If Not IsNull(DtaCuentas.Recordset("Apellido2")) Then
   Me.TxtApellido2.Text = DtaCuentas.Recordset("Apellido2")
  End If
  
  If Not IsNull(DtaCuentas.Recordset("Cedula")) Then
   Me.txtcedula.Text = DtaCuentas.Recordset("Cedula")
  End If
  
  If Not IsNull(DtaCuentas.Recordset("RUC")) Then
   Me.TxtRUC.Text = DtaCuentas.Recordset("RUC")
  End If
  
  If Not IsNull(DtaCuentas.Recordset("Telefono")) Then
   Me.TxtTelefono.Text = DtaCuentas.Recordset("Telefono")
  End If
  
  If Not IsNull(DtaCuentas.Recordset("Direccion")) Then
   Me.TxtDireccion.Text = DtaCuentas.Recordset("Direccion")
  End If
  
  If Not IsNull(DtaCuentas.Recordset("SubDivicion")) Then
   Me.CmbAgrupado.Text = DtaCuentas.Recordset("SubDivicion")
  Else
   Me.CmbAgrupado.Text = ""
  End If
  
  If Not IsNull(DtaCuentas.Recordset("TipoMoneda")) Then
   CmbTipoMoneda = DtaCuentas.Recordset("TipoMoneda")
  End If
  
  If Not IsNull(DtaCuentas.Recordset("CodCuentaImporta")) Then
   Me.TxtCuentaImporta.Text = DtaCuentas.Recordset("CodCuentaImporta")
  Else
   Me.TxtCuentaImporta.Text = ""
  End If
  
  
  If Not IsNull(DtaCuentas.Recordset("UbicacionReporte")) Then
    Me.CmbUbicacion.Text = DtaCuentas.Recordset("UbicacionReporte")
    Me.CmbUbicacionResultado.Text = DtaCuentas.Recordset("UbicacionReporte")
  Else
    Me.CmbUbicacion.Text = ""
    Me.CmbUbicacionResultado.Text = ""
  End If
  Me.LblSaldo.Caption = Format(DtaCuentas.Recordset("SaldoActual"), "##,##0.00")
 '/////Busco la Descripcion del Grupo/////////////////
   If Not IsNull(DtaCuentas.Recordset("CodGrupo")) Then
     CodGrupo = DtaCuentas.Recordset("CodGrupo")
   End If
  Criterio = "CodGrupo='" & CodGrupo & "'"
  If Me.DtaGrupoCuentas.Recordset.RecordCount > 0 Then Me.DtaGrupoCuentas.Recordset.MoveFirst
  Me.DtaGrupoCuentas.Recordset.Find Criterio
  If Not DtaGrupoCuentas.Recordset.EOF Then
    Me.DBGrupos.Text = DtaGrupoCuentas.Recordset("DescripcionGrupo")
  End If
  
   If Not IsNull(DtaCuentas.Recordset("KeyGrupo")) Then
     KeyGrupo = DtaCuentas.Recordset("KeyGrupo")
   End If
   
   Me.DtaConsulta.RecordSource = "SELECT KeyGrupo, CodGrupo, KeyGrupoSuperior, Child, DescripcionGrupo, Imagen1, Imagen2 From Grupos WHERE (KeyGrupo = '" & KeyGrupo & "')"
   Me.DtaConsulta.Refresh
   If Me.DtaConsulta.Recordset.EOF Then
     MsgBox "El grupo de Esta Cuenta no Existe,Corriga el Error!!!", vbCritical, "Sistema Contable Zeus"
     Me.TxtDescripcionGrupo.Text = ""
   Else
     Me.TxtDescripcionGrupo.Text = Me.DtaConsulta.Recordset("DescripcionGrupo")
   
   End If
   
  
  '//////////Muestro los Saldos de las cuentas/////////////////////
 Me.DtaSaldos.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento, Transacciones.DescripcionMovimiento,Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Debito*TCambio AS MDebito, TCambio*Credito AS MCredito From Transacciones Where (((Transacciones.CodCuentas) =  '" & Me.DBCliente.Text & "')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
 Me.DtaSaldos.Refresh
 
 
  Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento, Transacciones.DescripcionMovimiento, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, Debito*TCambio AS MDebito, TCambio*Credito AS MCredito, Transacciones.TCambio From Transacciones Where (((Transacciones.CodCuentas) =  '" & Me.DBCliente.Text & "')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
  Me.DtaConsulta.Refresh
  If DtaConsulta.Recordset.EOF Then
'     Me.CmbTipoMoneda.Enabled = True
     Me.CmbTipo.Enabled = True
  Else
     Me.CmbTipoMoneda.Enabled = False
     Me.CmbTipo.Enabled = False
  End If
  
  
  
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
  
''  Me.DBGCuentas.Columns(5).Visible = False
''  Me.DBGCuentas.Columns(6).Visible = False
'  Me.DBGCuentas.Columns(7).Caption = "Debito"
'  Me.DBGCuentas.Columns(8).Caption = "Credito"
'  Me.DBGCuentas.Columns(7).NumberFormat = "##,##0.00"
'  Me.DBGCuentas.Columns(8).NumberFormat = "##,##0.00"
'  Me.DBGCuentas.Columns(4).NumberFormat = "##,##0.000000"
  Me.LblSaldo.Caption = Format(Total1, "##,##0.00")
' Me.DBGCuentas.Columns(0).Visible = False


 Select Case Me.CmbTipo.Text
   Case "Cuentas x Cobrar": Me.ChkCentroCostos.Visible = True
   Case "Inventario": Me.ChkCentroCostos.Visible = True
   Case "Otros Activos": Me.ChkCentroCostos.Visible = True
   Case Else: Me.ChkCentroCostos.Visible = False
 End Select

 
  If Not IsNull(DtaCuentas.Recordset("CentroCostos")) Then
   If DtaCuentas.Recordset("CentroCostos") = 0 Then
      Me.ChkCentroCostos.Value = xtpUnchecked
   Else
       Me.ChkCentroCostos.Value = xtpChecked
   End If
  Else
  Me.ChkCentroCostos.Value = xtpUnchecked
      
 End If
 
 
 
End If
Exit Sub
TipoErrs:
ControlErrores

End Sub
