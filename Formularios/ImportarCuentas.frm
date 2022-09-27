VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmImportarCuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importacion de Cuentas"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   1815
      Begin VB.OptionButton OptExcel 
         Caption         =   "Archivos Excel"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   430
         Width           =   1455
      End
      Begin VB.OptionButton OptCns 
         Caption         =   "Archivos Cns"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   160
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   7440
      TabIndex        =   6
      Top             =   3360
      Width           =   1455
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5106
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Codigo Cuenta"
      Columns(0).DataField=   "CodigoCuenta"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "CodigoDepartamento"
      Columns(1).DataField=   "CodigoDepartamento"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "DescripcionCuenta"
      Columns(2).DataField=   "DescripcionCuenta"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "ClaseCuenta"
      Columns(3).DataField=   "ClaseCuenta"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "DescripcionClase"
      Columns(4).DataField=   "DescripcionClase"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   1085
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).DividerColor=   16315377
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
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
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   16315377
      RowDividerColor =   16315377
      RowSubDividerColor=   16315377
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=2,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(50)  =   "Named:id=33:Normal"
      _StyleDefs(51)  =   ":id=33,.parent=0"
      _StyleDefs(52)  =   "Named:id=34:Heading"
      _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(54)  =   ":id=34,.wraptext=-1"
      _StyleDefs(55)  =   "Named:id=35:Footing"
      _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(57)  =   "Named:id=36:Selected"
      _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=37:Caption"
      _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(61)  =   "Named:id=38:HighlightRow"
      _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(63)  =   "Named:id=39:EvenRow"
      _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(65)  =   "Named:id=40:OddRow"
      _StyleDefs(66)  =   ":id=40,.parent=33"
      _StyleDefs(67)  =   "Named:id=41:RecordSelector"
      _StyleDefs(68)  =   ":id=41,.parent=34"
      _StyleDefs(69)  =   "Named:id=42:FilterBar"
      _StyleDefs(70)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton CmdLeer 
      Caption         =   "Leer Registros"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton CmdAnexar 
      Caption         =   "Anexar Cuentas"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   4560
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblProgreso 
      Height          =   375
      Left            =   2160
      OleObjectBlob   =   "ImportarCuentas.frx":0000
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   4935
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblRegistros 
      Height          =   495
      Left            =   3600
      OleObjectBlob   =   "ImportarCuentas.frx":005E
      TabIndex        =   2
      Top             =   3360
      Width           =   2055
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   495
      Left            =   2160
      OleObjectBlob   =   "ImportarCuentas.frx":00BC
      TabIndex        =   3
      Top             =   3360
      Width           =   975
   End
   Begin MSAdodcLib.Adodc AdoAnexar 
      Height          =   375
      Left            =   1080
      Top             =   5880
      Width           =   5055
      _ExtentX        =   8916
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
      Caption         =   "AdoAnexar"
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
   Begin MSAdodcLib.Adodc AdoCuenta 
      Height          =   495
      Left            =   1560
      Top             =   6120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   873
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
      Caption         =   "AdoCuenta"
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
      Left            =   480
      Top             =   6000
      Width           =   4695
      _ExtentX        =   8281
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
   Begin MSAdodcLib.Adodc AdoImporta 
      Height          =   375
      Left            =   960
      Top             =   6600
      Width           =   4815
      _ExtentX        =   8493
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
      Caption         =   "AdoImporta"
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
   Begin MSAdodcLib.Adodc AdoRegistros 
      Height          =   375
      Left            =   1080
      Top             =   5760
      Width           =   4815
      _ExtentX        =   8493
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
      Caption         =   "AdoRegistros"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6600
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeSuiteControls.ProgressBar osProgress 
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   3960
      Visible         =   0   'False
      Width           =   6855
      _Version        =   786432
      _ExtentX        =   12091
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
End
Attribute VB_Name = "FrmImportarCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAnexar_Click()
 Dim KeyGrupo As String, DescripcionGrupo As String, CantidadRegistros As Double
 Dim CodigoCuenta As String, DescripcionCuenta As String
 Dim TipoCuenta As String, CodGrupo As String, CodTipoCuenta As String

 Me.AdoRegistros.Recordset.MoveLast
 CantidadRegistros = Me.AdoRegistros.Recordset.RecordCount
 Me.osProgress.Visible = True
 Me.LblProgreso.Visible = True
With Me.osProgress
 .Min = 0
 .Max = CantidadRegistros
 .Value = 0
 J = 1
 
    Me.TDBGrid.MoveFirst
    Do While Not Me.TDBGrid.EOF

'    Me.AdoRegistros.Refresh
'    Do While Not Me.AdoRegistros.Recordset.EOF
'        CodigoCuenta = Me.AdoRegistros.Recordset("CodigoCuenta")
'        DescripcionCuenta = Me.AdoRegistros.Recordset("DescripcionCuenta")
'        CodTipoCuenta = Me.AdoRegistros.Recordset("ClaseCuenta")
        
            CodigoCuenta = Me.TDBGrid.Columns(0)
            DescripcionCuenta = Me.TDBGrid.Columns(2)
            CodTipoCuenta = Me.TDBGrid.Columns(3)
        Me.LblProgreso.Caption = "Agregando la Cuenta: " & CodigoCuenta & " " & DescripcionCuenta
   
        Select Case CodTipoCuenta
        Case "01"
                 TipoCuenta = "Otros Activos"
                 KeyGrupoCuenta = "A"
        Case "02"
                 TipoCuenta = "Caja"
                 KeyGrupoCuenta = "A"
        Case "03"
                 TipoCuenta = "Cuentas x Cobrar"
                  KeyGrupoCuenta = "A"
        Case "04"
                 TipoCuenta = "Otros Activos"
                  KeyGrupoCuenta = "A"
        Case "05"
                 TipoCuenta = "Inventarios"
                  KeyGrupoCuenta = "A"
        Case "06"
                 TipoCuenta = "Activo Fijo"
                  KeyGrupoCuenta = "A"
        Case "07"
                 TipoCuenta = "Pasivo"
                  KeyGrupoCuenta = "B"
        Case "08"
                 TipoCuenta = "Capital"
                  KeyGrupoCuenta = "C"
        Case "09"
                 TipoCuenta = "Ingresos - Ventas"
                  KeyGrupoCuenta = "D"
        Case "10"
                  TipoCuenta = "Gastos"
                   KeyGrupoCuenta = "O"
        Case "11"
                  TipoCuenta = "Cuentas de Orden"
                   KeyGrupoCuenta = "P"
  
       End Select
   
   
     '//////////////////////////////////////////////////////
     '////BUSCO LA DESCRIPCION DEL GRUPO//////////////////
     '////////////////////////////////////////////////////////
   
        Me.AdoConsulta.RecordSource = "SELECT KeyGrupo, CodGrupo, KeyGrupoSuperior, Child, DescripcionGrupo, Imagen1, Imagen2 " & _
                                  "From Grupos WHERE (KeyGrupo = '" & KeyGrupoCuenta & "')"
        Me.AdoConsulta.Refresh
        If Not Me.AdoConsulta.Recordset.EOF Then
            DescripcionGrupo = Me.AdoConsulta.Recordset("DescripcionGrupo")
        Else
            MsgBox "Los Grupos no estan Correctos", vbCritical, "Sistema Contable"
            Exit Sub
        End If
    
       '//////////////////////////////////////////////////////////////
        '////////////////BUSCO SI EXISTE EL DEPARTAMENTO////////////////
        '/////////////////////////////////////////////////////////////////
        Me.AdoConsulta.RecordSource = "SELECT CodGrupo, DescripcionGrupo From GrupoCuentas WHERE (CodGrupo = '" & Me.AdoRegistros.Recordset("CodigoDepartamento") & "')"
        Me.AdoConsulta.Refresh
        If Not Me.AdoConsulta.Recordset.EOF Then
            CodGrupo = Me.AdoConsulta.Recordset("CodGrupo")
    
        End If
    
 
        '////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////BUSCO SI EXISTE LA CUENTA/////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////
 
        Me.AdoCuenta.RecordSource = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo " & _
        "From Cuentas WHERE     (CodCuentas = '" & CodigoCuenta & "')"
        Me.AdoCuenta.Refresh
        If Me.AdoCuenta.Recordset.EOF Then
 
            Me.AdoCuenta.Recordset.AddNew
            Me.AdoCuenta.Recordset("KeyGrupo") = KeyGrupoCuenta
            Me.AdoCuenta.Recordset("DescripcionGrupo") = DescripcionGrupo
            Me.AdoCuenta.Recordset("CodCuentas") = CodigoCuenta
            Me.AdoCuenta.Recordset("DescripcionCuentas") = DescripcionCuenta
            Me.AdoCuenta.Recordset("TipoCuenta") = TipoCuenta
            If CodGrupo <> "" Then
                Me.AdoCuenta.Recordset("CodGrupo") = CodGrupo
            End If
            Me.AdoCuenta.Recordset("TipoMoneda") = "Córdobas"
            Me.AdoCuenta.Recordset("SaldoActual") = 0#
            Me.AdoCuenta.Recordset.Update
        End If
    
        .Value = J
        J = J + 1
        
        
'        Me.AdoRegistros.Recordset.MoveNext
        Me.TDBGrid.MoveNext
    Loop
 End With
 
 
MsgBox "Se agregaron las cuentas!!", vbExclamationm, "Sistema Contable"



End Sub

Private Sub CmdLeer_Click()
On Error GoTo TipoErrs
Dim cadena As String
Dim Consecutivo As Integer, CuentaArchivo As String, CodigoArchivo As String
Dim CodigoCuenta As String, CodigoDepartamento As String, DescripcionArchivo As String
Dim DescripcionCuenta As String, Identificador As String
Dim Directorio As String, ClaseCuenta As String, DescripcionClase As String


If Me.OptCns.Value = True Then
 CommonDialog1.Filter = "Archivos Cns|*.Cns"
Else
 CommonDialog1.Filter = "Archivos de Excel|*.xls"
End If

Me.CommonDialog1.ShowOpen
Directorio = Me.CommonDialog1.FileName
If Directorio = "" Then
 Exit Sub
End If
'Como leer un fichero de texto desde Visual Basic

'Ejemplo de Programa

'1. Crear un nuevo proyecto en Visual basic, por defecto será Form1

'2. Añadir un control Text Box al formulario, fijar en propiedades "Multiline"
'en "TRUE" y en "ScrollBars" en 3-Both.

'3 Añadir el siguiente codigo al formulario en "Form_Load"

Me.AdoRegistros.Refresh
Do While Not Me.AdoRegistros.Recordset.EOF
  Me.AdoRegistros.Recordset.Delete
 Me.AdoRegistros.Recordset.MoveNext
Loop

Consecutivo = 0
crlf$ = Chr(13) & Chr(10)
Contador = 0
Open Directorio For Input As #1

Me.AdoImporta.RecordSource = "SELECT Identificador, CodigoCuenta, CodigoDepartamento, DescripcionCuenta, ClaseCuenta,DescripcionClase From RegistrosCuentas ORDER BY CodigoCuenta"
Me.AdoImporta.Refresh


   If Me.OptCns.Value = True Then
 
        While Not EOF(1)
        Salir = False
        Line Input #1, cadena
        
         Identificador = Mid(cadena, 1, 1)
         CuentaArchivo = Mid(cadena, 2, 16)
         CodigoDepartamento = Mid(cadena, 18, 3)
         DescripcionCuenta = Mid(cadena, 21, 35)
         ClaseCuenta = Mid(cadena, 56, 2)
        
        '/////////////QUITOS LOS ESPACIOS DEL CODIGO CUENTA///////////////
          Longitud = Len(CuentaArchivo)
          CodigoArchivo = ""
          CodigoCuenta = ""
         For i = 1 To Longitud
             CodigoArchivo = Mid(CuentaArchivo, i, 1)
             If CodigoArchivo <> " " Then
               CodigoCuenta = CodigoCuenta & CodigoArchivo
             
             End If
           
           Next
           
           
        '///////////////QUITOS LOS ESPACIOS DE LA DESCRIPCION//////////////////////
        '  Longitud = Len(DescripcionArchivo)
        '  CodigoArchivo = ""
        '  DescripcionCuenta = ""
        ' For I = 1 To Longitud
        '     CodigoArchivo = Mid(DescripcionArchivo, I, 1)
        '     If CodigoArchivo <> " " Then
        '       DescripcionCuenta = DescripcionCuenta & CodigoArchivo
        '
        '     End If
        '
        '   Next
           
         
        If Not CodigoCuenta = "                " Then
         
          Me.AdoRegistros.Recordset.AddNew
          AdoRegistros.Recordset("Identificador") = Identificador
          AdoRegistros.Recordset("CodigoCuenta") = CodigoCuenta
          AdoRegistros.Recordset("CodigoDepartamento") = CodigoDepartamento
          AdoRegistros.Recordset("DescripcionCuenta") = DescripcionCuenta
          AdoRegistros.Recordset("ClaseCuenta") = ClaseCuenta
          
          Select Case ClaseCuenta
               Case "01": DescripcionClase = "Otros Activos"
               Case "02": DescripcionClase = "Caja o Bancos"
               Case "03": DescripcionClase = "Ctas x Cob"
               Case "04": DescripcionClase = "Inversiones"
               Case "05": DescripcionClase = "Inventarios"
               Case "06": DescripcionClase = "Activo Fijo"
               Case "07": DescripcionClase = "Pasivos"
               Case "08": DescripcionClase = "Capital"
               Case "09": DescripcionClase = "Ingresos/Ventas"
               Case "10": DescripcionClase = "Gastos/Costos"
               Case "11": DescripcionClase = "OtrasCuentas"
          
          End Select
           AdoRegistros.Recordset("DescripcionClase") = DescripcionClase
           
          Me.AdoRegistros.Recordset.Update
          Contador = Contador + 1
          Me.LblRegistros.Caption = Contador
          DoEvents
         Else
          Anulado = Anulado + 1
         End If
        
        Wend
        Close #1

        Me.AdoImporta.RecordSource = "SELECT Identificador, CodigoCuenta, CodigoDepartamento, DescripcionCuenta, ClaseCuenta , DescripcionClase From RegistrosCuentas ORDER BY CodigoCuenta"
        Me.AdoImporta.Refresh
        Me.TDBGrid.DataSource = Me.AdoImporta
        
    Else
       

          Dim obj As New Class1
      
          Set Me.TDBGrid.DataSource = obj.Leer_Excel(Directorio, "Hoja1")
           Set obj = Nothing
          
          Me.AdoRegistros.Refresh
          Do While Not Me.AdoRegistros.Recordset.EOF
              Me.AdoRegistros.Recordset.Delete
             Me.AdoRegistros.Recordset.MoveNext
          Loop
          
          Me.TDBGrid.MoveFirst
          Do While Not Me.TDBGrid.EOF
            Identificador = "A"
            CodigoCuenta = Me.TDBGrid.Columns(0)
            CodigoDepartamento = Me.TDBGrid.Columns(1)
            DescripcionCuenta = Me.TDBGrid.Columns(2)
            ClaseCuenta = Me.TDBGrid.Columns(3)
              Select Case ClaseCuenta
                  Case "01": DescripcionClase = "Otros Activos"
                  Case "02": DescripcionClase = "Caja o Bancos"
                  Case "03": DescripcionClase = "Ctas x Cob"
                  Case "04": DescripcionClase = "Inversiones"
                  Case "05": DescripcionClase = "Inventarios"
                  Case "06": DescripcionClase = "Activo Fijo"
                  Case "07": DescripcionClase = "Pasivos"
                  Case "08": DescripcionClase = "Capital"
                  Case "09": DescripcionClase = "Ingresos/Ventas"
                  Case "10": DescripcionClase = "Gastos/Costos"
                  Case "11": DescripcionClase = "OtrasCuentas"
             
             End Select
             
                  If CodigoCuenta <> "" Then
                    Me.AdoRegistros.Recordset.AddNew
                    AdoRegistros.Recordset("Identificador") = Identificador
                    AdoRegistros.Recordset("CodigoCuenta") = CodigoCuenta
                    AdoRegistros.Recordset("CodigoDepartamento") = CodigoDepartamento
                    AdoRegistros.Recordset("DescripcionCuenta") = DescripcionCuenta
                    AdoRegistros.Recordset("ClaseCuenta") = ClaseCuenta
                    AdoRegistros.Recordset("DescripcionClase") = DescripcionClase
                    Me.AdoRegistros.Recordset.Update
                  End If
             
             Me.TDBGrid.MoveNext
          Loop
          
    
    End If

MsgBox "Proceso Terminado", vbExclamation, "Sistema de Enlace"
If Not Anulado = 0 Then
 MsgBox "Se quitaron " & Anulado & " Transacciones anuladas"
End If

Salir = True
Exit Sub
TipoErrs:
MsgBox err.Description
Salir = True
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd
 Me.TDBGrid.EvenRowStyle.BackColor = &H80FFFF
 Me.TDBGrid.OddRowStyle.BackColor = &HC0FFFF
 Me.TDBGrid.AlternatingRowStyle = True
 
With Me.AdoCuenta
   .ConnectionString = Conexion
End With

With Me.AdoImporta
   .ConnectionString = Conexion
End With

With Me.AdoAnexar
   .ConnectionString = Conexion
End With

With Me.AdoConsulta
   .ConnectionString = Conexion
End With

With Me.AdoRegistros
   .ConnectionString = Conexion
   .RecordSource = "RegistrosCuentas"
   .Refresh
End With
 
End Sub
