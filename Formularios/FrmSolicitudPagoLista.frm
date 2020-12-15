VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FrmSolicitudPagoLista 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Pagos"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   14385
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   375
      Left            =   12840
      TabIndex        =   11
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   82837505
      CurrentDate     =   44168
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROCESAR"
      Height          =   375
      Left            =   12840
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc AdoConsulta 
      Height          =   375
      Left            =   1080
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
      Caption         =   "AdoCosulta"
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
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   12840
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
      Begin VB.OptionButton OptProcesados 
         Caption         =   "Procesados"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton OptAnulados 
         Caption         =   "Anulados"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton OptActivos 
         Caption         =   "Activos"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "SALIR"
      Height          =   375
      Left            =   12840
      TabIndex        =   4
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "ANULAR"
      Height          =   375
      Left            =   12840
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "EDITAR"
      Height          =   375
      Left            =   12840
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "NUEVO"
      Height          =   375
      Left            =   12840
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin TrueOleDBGrid80.TDBGrid DBGTransacciones 
      Bindings        =   "FrmSolicitudPagoLista.frx":0000
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   10821
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
      Splits(0).Caption=   "Lista de Solicitud de Pagos"
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
      AllowUpdate     =   0   'False
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
   Begin MSAdodcLib.Adodc DtaIndice 
      Height          =   375
      Left            =   1440
      Top             =   5880
      Visible         =   0   'False
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
   Begin MSAdodcLib.Adodc AdoDetalleCheque 
      Height          =   375
      Left            =   1080
      Top             =   7680
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
      Caption         =   "AdoDetalleCheque"
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
Attribute VB_Name = "FrmSolicitudPagoLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ew As cls_NumEnglishWord
Private sw As cls_NumSpanishWord

Private Sub CmdEditar_Click()
  Dim Fecha As Date, NumeroSolicitud As Double, Nperiodo As Double, SQL As String
  Dim TotalDebito As Double, TotalCredito As Double, Monto As Double
  
    If Me.OptActivos.Value = False Then
       Me.OptActivos.Value = True
       ActualizaGrid
    End If
    
    Fecha = Me.DBGTransacciones.Columns("FechaTransaccion")
    NumeroSolicitud = Me.DBGTransacciones.Columns("NumeroSolicitud")
    
    QUIEN = "Editar"
    FrmSolicitudPagos.NumeroSolicitud = NumeroSolicitud
    
    Me.AdoConsulta.RecordSource = "SELECT IndiceSolicitudPago.* From IndiceSolicitudPago  " & _
                                  "WHERE  (FechaTransaccion = CONVERT(DATETIME, '" & Format(Fecha, "yyyy-mm-dd") & "', 102)) AND (NumeroMovimiento = " & NumeroSolicitud & ")"
    Me.AdoConsulta.Refresh
    If Not Me.AdoConsulta.Recordset.EOF Then
      FrmSolicitudPagos.TxtFecha.Value = Fecha
      FrmSolicitudPagos.DBCodigo.Text = Me.AdoConsulta.Recordset("CuentaBanco")
      FrmSolicitudPagos.DBCodigo_ItemChange
      FrmSolicitudPagos.TxtMemo.Text = Me.AdoConsulta.Recordset("Concepto")
      FrmSolicitudPagos.CmbMoneda.Text = Me.AdoConsulta.Recordset("TipoMoneda")
      FrmSolicitudPagos.ChkCheque.Value = Me.AdoConsulta.Recordset("ImprimeCheque")
      FrmSolicitudPagos.TxtSubTotal.Text = Format(Me.AdoConsulta.Recordset("SubTotal"), "##,##0.00")
      FrmSolicitudPagos.TxtIVa.Text = Format(Me.AdoConsulta.Recordset("MontoIva"), "##,##0.00")
      FrmSolicitudPagos.TxtRetenciones.Text = Format(Me.AdoConsulta.Recordset("MontoRetenciones"), "##,##0.00")
      FrmSolicitudPagos.TxtMonto.Text = Format(Me.AdoConsulta.Recordset("MontoSolicitud"), "##,##0.00")
      FrmSolicitudPagos.TxtNombre.Text = Format(Me.AdoConsulta.Recordset("Beneficiario"), "##,##0.00")
      
      Monto = Me.AdoConsulta.Recordset("MontoSolicitud")
      
      Nperiodo = Me.AdoConsulta.Recordset("NPeriodo")

       If Me.AdoConsulta.Recordset("TipoMoneda") = "Dólares" Then
         FrmSolicitudPagos.TxtLetras.Text = sw.ConvertCurrencyToSpanish(CDbl(FrmSolicitudPagos.TxtMonto.Text), "Dólares")
        ElseIf Me.AdoConsulta.Recordset("TipoMoneda") = "Córdobas" Then
         FrmSolicitudPagos.TxtLetras.Text = sw.ConvertCurrencyToSpanish(CDbl(FrmSolicitudPagos.TxtMonto.Text), "Córdobas")
        End If

      If Me.AdoConsulta.Recordset("Anticipo") = 1 Then
        FrmSolicitudPagos.OptAnticipo.Value = True
      Else
        FrmSolicitudPagos.OptCancelacion.Value = True
      End If
      
       If Me.AdoConsulta.Recordset("Retencion1") = "Verdadero" Then
         FrmSolicitudPagos.Chk1.Value = 1
       Else
         FrmSolicitudPagos.Chk1.Value = 0
       End If
       
       If Me.AdoConsulta.Recordset("Retencion2") = "Verdadero" Then
         FrmSolicitudPagos.Chk2.Value = 1
       Else
         FrmSolicitudPagos.Chk2.Value = 0
       End If
       
       If Me.AdoConsulta.Recordset("Retencion3") = "Verdadero" Then
         FrmSolicitudPagos.Chk3.Value = 1
       Else
         FrmSolicitudPagos.Chk3.Value = 0
       End If
       
       If Me.AdoConsulta.Recordset("Retencion4") = "Verdadero" Then
         FrmSolicitudPagos.Chk7.Value = 1
       Else
         FrmSolicitudPagos.Chk7.Value = 0
       End If
       
       If Me.AdoConsulta.Recordset("Retencion5") = "Verdadero" Then
          FrmSolicitudPagos.Chk10.Value = 1
       Else
           FrmSolicitudPagos.Chk10.Value = 0
       End If
       FrmSolicitudPagos.TxtNTransacciones.Text = NumeroSolicitud
       
       
       '//////////////////////////////////////////////////////////////////////////////////////////////
       '////////////////////////CONSULTO TRANSACCIONES DE PAGO //////////////////////////////////7////
       '///////////////////////////////////////////////////////////////////////////////////////////////
       SQL = "SELECT     TransaccionesSolicitudPago.CodCuentas, TransaccionesSolicitudPago.NombreCuenta, TransaccionesSolicitudPago.VoucherNo, TransaccionesSolicitudPago.DescripcionMovimiento, " & _
       "TransaccionesSolicitudPago.FacturaNo, TransaccionesSolicitudPago.ChequeNo, TransaccionesSolicitudPago.Clave, TransaccionesSolicitudPago.TCambio, TransaccionesSolicitudPago.Debito, TransaccionesSolicitudPago.Credito, " & _
       "TransaccionesSolicitudPago.FechaTransaccion, TransaccionesSolicitudPago.NPeriodo, TransaccionesSolicitudPago.NTransaccion, TransaccionesSolicitudPago.Fuente, TransaccionesSolicitudPago.FechaTasas, " & _
       "TransaccionesSolicitudPago.NumeroMovimiento, Periodos.Periodo, TransaccionesSolicitudPago.FechaDescuento, TransaccionesSolicitudPago.DescuentoDisponible, " & _
       "TransaccionesSolicitudPago.FechaVence,TransaccionesSolicitudPago.CodCuentaProveedor,TransaccionesSolicitudPago.TipoFactura,TransaccionesSolicitudPago.NTransaccion " & _
       "FROM  Periodos INNER JOIN " & _
       "TransaccionesSolicitudPago ON Periodos.NPeriodo = TransaccionesSolicitudPago.NPeriodo " & _
       "WHERE  (TransaccionesSolicitudPago.FechaTransaccion = CONVERT(DATETIME, '" & Format(Fecha, "yyyy-mm-dd") & "', 102)) AND (TransaccionesSolicitudPago.NumeroMovimiento = " & NumeroSolicitud & ") AND (TransaccionesSolicitudPago.NPeriodo = " & Nperiodo & ")" & _
       "ORDER BY TransaccionesSolicitudPago.NTransaccion "
       
        FrmSolicitudPagos.DtaTransacciones.RecordSource = SQL
        FrmSolicitudPagos.DtaTransacciones.Refresh
        
        FrmSolicitudPagos.DtaBancos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta From Cuentas WHERE (TipoCuenta = 'Caja') OR (TipoCuenta = N'Bancos') ORDER BY Cuentas.CodCuentas"
        FrmSolicitudPagos.DtaBancos.Refresh
        FrmSolicitudPagos.DBCodigo.ListField = "CodCuentas"
        
    
        SQL = "SELECT  MAX(TransaccionesSolicitudPago.CodCuentas) AS CodCuentas, SUM(TransaccionesSolicitudPago.Debito) AS Debito, SUM(TransaccionesSolicitudPago.Credito) AS Credito, SUM(TransaccionesSolicitudPago.DebitoD) AS DebitoD, SUM(TransaccionesSolicitudPago.CreditoD) AS CreditoD FROM  Periodos INNER JOIN  TransaccionesSolicitudPago ON Periodos.NPeriodo = TransaccionesSolicitudPago.NPeriodo  " & _
              "WHERE  (TransaccionesSolicitudPago.FechaTransaccion = CONVERT(DATETIME, '" & Format(Fecha, "yyyy-mm-dd") & "', 102)) AND (TransaccionesSolicitudPago.NumeroMovimiento = " & NumeroSolicitud & ") AND (TransaccionesSolicitudPago.NPeriodo = " & Nperiodo & ")"
        FrmSolicitudPagos.AdoBuscar.RecordSource = SQL
        FrmSolicitudPagos.AdoBuscar.Refresh
        If Not FrmSolicitudPagos.AdoBuscar.Recordset.EOF Then
          If Not IsNull(FrmSolicitudPagos.AdoBuscar.Recordset("Debito")) Then
          TotalDebito = FrmSolicitudPagos.AdoBuscar.Recordset("Debito")
          End If
          If Not IsNull(FrmSolicitudPagos.AdoBuscar.Recordset("Credito")) Then
          TotalCredito = FrmSolicitudPagos.AdoBuscar.Recordset("Credito") + Monto
          End If
          FrmSolicitudPagos.TxtDebito.Text = Format(TotalDebito, "##,##0.00")
          FrmSolicitudPagos.TxtCredito.Text = Format(TotalCredito, "##,##0.00")
          FrmSolicitudPagos.TxtDiferencia.Text = Format(TotalDebito - TotalCredito, "##,##0.00")
        End If
         
        
          FrmSolicitudPagos.DBGTransacciones.Columns("CodCuentas").Button = True
          FrmSolicitudPagos.DBGTransacciones.Columns("NombreCuenta").Locked = True
          FrmSolicitudPagos.DBGTransacciones.Columns("NombreCuenta").Locked = True
          FrmSolicitudPagos.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
          FrmSolicitudPagos.DBGTransacciones.Columns(6).Button = True
          FrmSolicitudPagos.DBGTransacciones.Columns(6).Locked = True
          FrmSolicitudPagos.DBGTransacciones.Columns(0).Width = 1500
          FrmSolicitudPagos.DBGTransacciones.Columns(2).Width = 1000
          FrmSolicitudPagos.DBGTransacciones.Columns(3).Caption = "Descripcion"
          FrmSolicitudPagos.DBGTransacciones.Columns(4).Width = 1000
          FrmSolicitudPagos.DBGTransacciones.Columns(4).Button = True
          FrmSolicitudPagos.DBGTransacciones.Columns(5).Width = 1000
          FrmSolicitudPagos.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
          FrmSolicitudPagos.DBGTransacciones.Columns(6).Width = 800
          FrmSolicitudPagos.DBGTransacciones.Columns(7).Caption = "Tasa Cambio"
          FrmSolicitudPagos.DBGTransacciones.Columns(7).Locked = True
          FrmSolicitudPagos.DBGTransacciones.Columns(7).NumberFormat = "##,##0.000000"
          FrmSolicitudPagos.DBGTransacciones.Columns(7).Width = 1200
          FrmSolicitudPagos.DBGTransacciones.Columns(8).Width = 1200
          FrmSolicitudPagos.DBGTransacciones.Columns(8).NumberFormat = "##,##0.00"
          FrmSolicitudPagos.DBGTransacciones.Columns(9).Width = 1200
          FrmSolicitudPagos.DBGTransacciones.Columns(9).NumberFormat = "##,##0.00"
          FrmSolicitudPagos.DBGTransacciones.Columns(10).Visible = False
          FrmSolicitudPagos.DBGTransacciones.Columns(11).Visible = False
          FrmSolicitudPagos.DBGTransacciones.Columns(12).Visible = False
          FrmSolicitudPagos.DBGTransacciones.Columns(13).Visible = False
          FrmSolicitudPagos.DBGTransacciones.Columns(14).Visible = False
          FrmSolicitudPagos.DBGTransacciones.Columns(15).Visible = False
          FrmSolicitudPagos.DBGTransacciones.Columns(16).Visible = False
          FrmSolicitudPagos.DBGTransacciones.Columns(17).Visible = False
          FrmSolicitudPagos.DBGTransacciones.Columns(18).Visible = False
          FrmSolicitudPagos.DBGTransacciones.Columns(19).Visible = False
          FrmSolicitudPagos.DBGTransacciones.Columns(20).Visible = False
          FrmSolicitudPagos.DBGTransacciones.Columns(21).Visible = False
          FrmSolicitudPagos.DBGTransacciones.Columns(22).Visible = False
          FrmSolicitudPagos.DBGTransacciones.Columns(7).Locked = True 'columna tasa de cambio
       
            
         
      
    End If
    

    FrmSolicitudPagos.Show 1
    FrmSolicitudPagos.DtaIndice.Refresh
End Sub

Private Sub CmdEliminar_Click()
  Dim Fecha As Date, NumeroSolicitud As Double, Nperiodo As Double, SQL As String
  Dim TotalDebito As Double, TotalCredito As Double, Monto As Double, Respuesta As Double
  
    If MsgBox("¿Esta Seguro de Anular la solicitud?", vbYesNo, "Zeus Contable") = vbNo Then
      Exit Sub
    End If
  
    If Me.OptActivos.Value = False Then
       Me.OptActivos.Value = True
       ActualizaGrid
    End If
    
    Fecha = Me.DBGTransacciones.Columns("FechaTransaccion")
    NumeroSolicitud = Me.DBGTransacciones.Columns("NumeroSolicitud")
    
    Me.AdoConsulta.RecordSource = "SELECT IndiceSolicitudPago.* From IndiceSolicitudPago  " & _
                                  "WHERE  (FechaTransaccion = CONVERT(DATETIME, '" & Format(Fecha, "yyyy-mm-dd") & "', 102)) AND (NumeroMovimiento = " & NumeroSolicitud & ")"
    Me.AdoConsulta.Refresh
    If Not Me.AdoConsulta.Recordset.EOF Then
      Me.AdoConsulta.Recordset("DescripcionMovimiento") = "****ANULADO****·"
      Me.AdoConsulta.Recordset("Activo") = 0
      Me.AdoConsulta.Recordset("Anulado") = 1
      Me.AdoConsulta.Recordset.Update
    End If
    
 
        ActualizaGrid
        
End Sub

Private Sub CmdNuevo_Click()
QUIEN = "Nuevo"
FrmSolicitudPagos.Show 1
ActualizaGrid

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()
  Dim NumeroTransaccion As Double
  Dim mes As Double, Año As Double, FechaIni As Date, FechaFin As Date
  Dim Periodo As Double, NumeroPeriodo As Double, EstadoPeriodo As String, SQL As String, TipoMoneda As String
  Dim Retencion1 As Boolean, Retencion2 As Boolean, Retencion3 As Boolean, Retencion4 As Boolean, Retencion5 As Boolean
  Dim CuentaContable As String, NumeroSolicitud As Double, Reg As Double, MonedaNomina As String, Resultado As Double
  Dim CuentaBanco As String, CuentaIva As String, CuentaRetencion1 As String, CuentaRetencion2 As String, CuentaRetencion3 As String, CuentaRetencion4 As String, CuentaRetencion5 As String
  Dim MontoBanco As Double, MontoRetencion1 As Double, MontoRetencion2 As Double, MontoRetencion3 As Double, MontoRetencion4 As Double, MontoRetencion5 As Double
  Dim DescripcionCuenta As String, DescripcionMovimiento As String, TipoMovimiento As String, NumeroFactura As String, Debito As Double, Credito As Double
  Dim TasaCambio As Double, FechaFactura As Date, Descuento As Double, VoucherNo As String, MontoCheque As Double, SubTotal As Double, MontoIva As Double, MontoRetenciones As Double
  Dim Iva As Boolean, ConceptoSolicitud As String, MontoIva2 As Double, MontoRetenciones2 As Double

  
    If MsgBox("¿Esta Seguro de querer PROCESAR?", vbYesNo, "Zeus Contable") = vbNo Then
      Exit Sub
    End If
    
    Reg = 1
    '///////////////////////////////////////////////////////////////////////////////////////////////////
    '/////////////////////////////////CONSULTO EL INDICE SOLICUTID PARA CREAR EL INDICE DEL CHEQUE //////
    '///////////////////////////////////////////////////////////////////////////////////////////////////

    Fecha = Me.DBGTransacciones.Columns("FechaTransaccion")
    NumeroSolicitud = Me.DBGTransacciones.Columns("NumeroSolicitud")
    DescripcionMovimiento = "Registro de Solcitud de Cheque No. " & NumeroSolicitud
    
    Me.AdoConsulta.RecordSource = "SELECT IndiceSolicitudPago.* From IndiceSolicitudPago  " & _
                                  "WHERE  (FechaTransaccion = CONVERT(DATETIME, '" & Format(Fecha, "yyyy-mm-dd") & "', 102)) AND (NumeroMovimiento = " & NumeroSolicitud & ")"
    Me.AdoConsulta.Refresh
    If Not Me.AdoConsulta.Recordset.EOF Then
    
       MonedaNomina = Me.AdoConsulta.Recordset("TipoMoneda")
       '///////////////////////////////////////////////////////////////AGREGO EL INDICIE DE CHEQUE /////////////////////////////////////////////////
        If Reg = 1 Then
        
              
                 mes = Month(Me.DTPFecha.Value)
                 Año = Year(Me.DTPFecha.Value)
                 FechaIni = CDate("1/" & Month(Me.DTPFecha.Value) & "/" & Year(Me.DTPFecha.Value))
                 
                        TipoMoneda = Me.AdoConsulta.Recordset("TipoMoneda")
                        Retencion1 = Me.AdoConsulta.Recordset("Retencion1")  'Retencion 1%
                        Retencion2 = Me.AdoConsulta.Recordset("Retencion2")  'Retencion 2%
                        Retencion3 = Me.AdoConsulta.Recordset("Retencion3")  'Retencion 3%
                        Retencion4 = Me.AdoConsulta.Recordset("Retencion4")  'Reteciion 4%
                        Retencion5 = Me.AdoConsulta.Recordset("Retencion5")  'Retencion 10%
                        Iva = Me.AdoConsulta.Recordset("Iva")
                        
                        If Not IsNull(Me.AdoConsulta.Recordset("Concepto")) Then
                          ConceptoSolicitud = Me.AdoConsulta.Recordset("Concepto")
                        End If
                        
                        If Not IsNull(Me.AdoConsulta.Recordset("MontoSolicitud")) Then
                          MontoCheque = Me.AdoConsulta.Recordset("MontoSolicitud")
                        Else
                          MontoCheque = 0
                        End If
                        If Not IsNull(Me.AdoConsulta.Recordset("SubTotal")) Then
                          SubTotal = Me.AdoConsulta.Recordset("SubTotal")
                        Else
                          SubTotal = 0
                        End If
                        If Not IsNull(Me.AdoConsulta.Recordset("MontoIva")) Then
                          MontoIva = Me.AdoConsulta.Recordset("MontoIva")
                        Else
                          MontoIva = 0
                        End If
                        If Not IsNull(Me.AdoConsulta.Recordset("MontoRetenciones")) Then
                          MontoRetenciones = Me.AdoConsulta.Recordset("MontoRetenciones")
                        Else
                          MontoRetenciones = 0
                        End If
                        
                        If Not IsNull(Me.AdoConsulta.Recordset("CuentaBanco")) Then
                          CuentaBanco = Me.AdoConsulta.Recordset("CuentaBanco")
                        Else
                          CuentaBanco = 0
                        End If
                        
                       
                       If Not IsNull(Me.AdoConsulta.Recordset("CtaIva")) Then
                          CuentaIva = Me.AdoConsulta.Recordset("CtaIva")
                        Else
                          CuentaIva = 0
                        End If
                        
                        If Not IsNull(Me.AdoConsulta.Recordset("CtaRetencion1")) Then
                          CuentaRetencion1 = Me.AdoConsulta.Recordset("CtaRetencion1")
                        Else
                          CuentaRetencion1 = 0
                        End If
                        
                        If Not IsNull(Me.AdoConsulta.Recordset("CtaRetencion2")) Then
                          CuentaRetencion2 = Me.AdoConsulta.Recordset("CtaRetencion2")
                        Else
                          CuentaRetencion2 = 0
                        End If
                        
                        If Not IsNull(Me.AdoConsulta.Recordset("CtaRetencion3")) Then
                          CuentaRetencion3 = Me.AdoConsulta.Recordset("CtaRetencion3")
                        Else
                          CuentaRetencion3 = 0
                        End If
                        
                       If Not IsNull(Me.AdoConsulta.Recordset("CtaRetencion4")) Then
                          CuentaRetencion4 = Me.AdoConsulta.Recordset("CtaRetencion4")
                        Else
                          CuentaRetencion4 = 0
                        End If
                        
                        If Not IsNull(Me.AdoConsulta.Recordset("CtaRetencion5")) Then
                          CuentaRetencion5 = Me.AdoConsulta.Recordset("CtaRetencion5")
                        Else
                          CuentaRetencion5 = 0
                        End If
                        
                        
                 FechaFin = DateSerial(Año, mes + 1, 1 - 1)

                 
                 Me.AdoConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "'))"
                 Me.AdoConsulta.Refresh
                 If Not Me.AdoConsulta.Recordset.EOF Then
                     Periodo = Me.AdoConsulta.Recordset("Periodo")
                     NumeroPeriodo = Me.AdoConsulta.Recordset("NPeriodo")
                     EstadoPeriodo = Me.AdoConsulta.Recordset("EstadoPeriodo")
                 End If
        
        
                   If EstadoPeriodo = "A" Then
        
                        Me.AdoConsulta.Recordset("NTransacciones") = Me.AdoConsulta.Recordset("NTransacciones") + 1
                        Me.AdoConsulta.Recordset.Update
                        NumeroTransaccion = Me.AdoConsulta.Recordset("NTransacciones")
                        
                        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '/////////////////////////////////////////AGREGO EL DETALLE DE LOS CHEQUES /////////////////////////////////
                        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////

                      
                        
                        '////////////////////////////////////////////////////////////////
                        '////////AGREGO LOS INDICES DE TRANSACCIONES//////
                        '///////////////////////////////////////////////////////////////
'                        MonedaNomina = "Córdobas"
                        Resultado = GrabaEncabezado(NumeroPeriodo, NumeroTransaccion, Format(Me.DTPFecha.Value, "yyyy-mm-dd"), "Solicitud  de Cheque", "CHEQUE", TipoMoneda)
                        Reg = 2
                        
                        
                     
                        
     

                               
                        '//////////////////////////////////////////////////////////////////////////////////////////////
                        '////////////////////////CONSULTO TRANSACCIONES DE PAGO //////////////////////////////////7////
                        '///////////////////////////////////////////////////////////////////////////////////////////////
                        SQL = "SELECT     TransaccionesSolicitudPago.CodCuentas, TransaccionesSolicitudPago.NombreCuenta, TransaccionesSolicitudPago.VoucherNo, TransaccionesSolicitudPago.DescripcionMovimiento, " & _
                        "TransaccionesSolicitudPago.FacturaNo, TransaccionesSolicitudPago.ChequeNo, TransaccionesSolicitudPago.Clave, TransaccionesSolicitudPago.TCambio, TransaccionesSolicitudPago.Debito, TransaccionesSolicitudPago.Credito, " & _
                        "TransaccionesSolicitudPago.FechaTransaccion, TransaccionesSolicitudPago.NPeriodo, TransaccionesSolicitudPago.NTransaccion, TransaccionesSolicitudPago.Fuente, TransaccionesSolicitudPago.FechaTasas, " & _
                        "TransaccionesSolicitudPago.NumeroMovimiento, Periodos.Periodo, TransaccionesSolicitudPago.FechaDescuento, TransaccionesSolicitudPago.DescuentoDisponible, " & _
                        "TransaccionesSolicitudPago.FechaVence,TransaccionesSolicitudPago.CodCuentaProveedor,TransaccionesSolicitudPago.TipoFactura,TransaccionesSolicitudPago.NTransaccion " & _
                        "FROM  Periodos INNER JOIN " & _
                        "TransaccionesSolicitudPago ON Periodos.NPeriodo = TransaccionesSolicitudPago.NPeriodo " & _
                        "WHERE  (TransaccionesSolicitudPago.FechaTransaccion = CONVERT(DATETIME, '" & Format(Fecha, "yyyy-mm-dd") & "', 102)) AND (TransaccionesSolicitudPago.NumeroMovimiento = " & NumeroSolicitud & ") " & _
                        "ORDER BY TransaccionesSolicitudPago.NTransaccion "
                        Me.AdoDetalleCheque.RecordSource = SQL
                        Me.AdoDetalleCheque.Refresh
                        Do While Not Me.AdoDetalleCheque.Recordset.EOF
                                
                                CuentaContable = Me.AdoDetalleCheque.Recordset("CodCuentas")
                                DescripcionCuenta = Me.AdoDetalleCheque.Recordset("NombreCuenta")
                                DescripcionMovimiento = Me.AdoDetalleCheque.Recordset("DescripcionMovimiento")
                                TipoMovimiento = Me.AdoDetalleCheque.Recordset("Clave")
                                NumeroFactura = Me.AdoDetalleCheque.Recordset("FacturaNo")
                                Debito = Me.AdoDetalleCheque.Recordset("Debito")
                                Credito = Me.AdoDetalleCheque.Recordset("Credito")
                                TasaCambio = Me.AdoDetalleCheque.Recordset("TCambio")
                                FechaFactura = Me.AdoDetalleCheque.Recordset("FechaDescuento")
                                VoucherNo = Me.AdoDetalleCheque.Recordset("VoucherNo")
                                Descuento = 0
                                
                                '///////////////////////////////VERIFICO EL MONTO DEL DEBITO ////////////////////////
                                If CDbl(Debito) <> 0 Then
                                   '//////////////////CALCULO SI E 100% ES DEBITO ///////////////////
                                   If CDbl(Format(MontoCheque / Debito, "####0")) = 1 Then
                                      '////////////////////SIGNIFICA QUE ES UNA LINEA /////////////////
                                       Debito = SubTotal
                                       Resultado = GrabaDetalleCheque(CuentaContable, Me.DTPFecha.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, TipoMovimiento, TasaCambio, Debito, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CuentaContable, DescripcionCuenta, VoucherNo)
                                   
                                   Else
                                   
                                   End If
                                   
                                End If
                                
                               

                                
'

                        
                           Me.AdoDetalleCheque.Recordset.MoveNext
                        Loop
                        
                        '////////////////////////////////////////////////////////////////////////////////////////
                        '///////////////////////AGREGO LAS CUENTAS DE IVA /////////////////////////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////
                        If Iva = True Then
                          Debito = MontoIva
                          Credito = 0
                          TipoMovimiento = "Debito"
                          Resultado = GrabaDetalleCheque(CuentaIva, Me.DTPFecha.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, TipoMovimiento, TasaCambio, Debito, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CuentaContable, DescripcionCuenta, VoucherNo)
                        End If
                        
                        '////////////////////////////////////////////////////////////////////////////////////////
                        '///////////////////////AGREGO LAS CUENTAS DE 1%/////////////////////////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////
                        If Retencion1 = True Then
                          Debito = 0
                          Credito = SubTotal * 0.01
                          TipoMovimiento = "Credito"
                          Resultado = GrabaDetalleCheque(CuentaRetencion1, Me.DTPFecha.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, TipoMovimiento, TasaCambio, Debito, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CuentaContable, DescripcionCuenta, VoucherNo)
                        End If
                        
                        '////////////////////////////////////////////////////////////////////////////////////////
                        '///////////////////////AGREGO LAS CUENTAS DE 2%/////////////////////////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////
                        If Retencion2 = True Then
                          Debito = 0
                          Credito = SubTotal * 0.02
                          TipoMovimiento = "Credito"
                          Resultado = GrabaDetalleCheque(CuentaRetencion2, Me.DTPFecha.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, TipoMovimiento, TasaCambio, Debito, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CuentaContable, DescripcionCuenta, VoucherNo)
                        End If
                        
                        '////////////////////////////////////////////////////////////////////////////////////////
                        '///////////////////////AGREGO LAS CUENTAS DE 3%/////////////////////////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////
                        If Retencion3 = True Then
                          Debito = 0
                          Credito = SubTotal * 0.03
                          TipoMovimiento = "Credito"
                          Resultado = GrabaDetalleCheque(CuentaRetencion3, Me.DTPFecha.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, TipoMovimiento, TasaCambio, Debito, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CuentaContable, DescripcionCuenta, VoucherNo)
                        End If
                        
                        
                        '////////////////////////////////////////////////////////////////////////////////////////
                        '///////////////////////AGREGO LAS CUENTAS DE 7%/////////////////////////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////
                        If Retencion4 = True Then
                          Debito = 0
                          Credito = SubTotal * 0.07
                          TipoMovimiento = "Credito"
                          Resultado = GrabaDetalleCheque(CuentaRetencion4, Me.DTPFecha.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, TipoMovimiento, TasaCambio, Debito, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CuentaContable, DescripcionCuenta, VoucherNo)
                        End If
                        
                                                '////////////////////////////////////////////////////////////////////////////////////////
                        '///////////////////////AGREGO LAS CUENTAS DE 7%/////////////////////////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////
                        If Retencion5 = True Then
                          Debito = 0
                          Credito = SubTotal * 0.1
                          TipoMovimiento = "Credito"
                          Resultado = GrabaDetalleCheque(CuentaRetencion5, Me.DTPFecha.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, TipoMovimiento, TasaCambio, Debito, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CuentaContable, DescripcionCuenta, VoucherNo)
                        End If
                        
                        '///////////////////////////////AGREGO LA CUENTA DE BANCO //////////////////
                        Debito = 0
                        Credito = MontoCheque
                        TipoMovimiento = "Credito"
                        Resultado = GrabaDetalleCheque(CuentaBanco, Me.DTPFecha.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, TipoMovimiento, TasaCambio, Debito, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CuentaContable, DescripcionCuenta, VoucherNo)
                        
                        
                        
                        
                   Else
                   
                       MsgBox "El Periodo Esta Cerrado o Bloqueado", vbCritical, "Zeus Contable"
                       Exit Sub
                   End If
         End If
    
    
    End If
    
    
    
End Sub

Private Sub Form_Initialize()
    Set ew = New cls_NumEnglishWord
    Set sw = New cls_NumSpanishWord
End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd

 Me.DBGTransacciones.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DBGTransacciones.OddRowStyle.BackColor = &H80000005
 Me.DBGTransacciones.AlternatingRowStyle = True

ActualizaGrid


Me.AdoConsulta.ConnectionString = Conexion

Me.AdoDetalleCheque.ConnectionString = Conexion


End Sub

Public Sub ActualizaGrid()
 If Me.OptTodos.Value = True Then
    Me.DtaIndice.ConnectionString = Conexion
    Me.DtaIndice.RecordSource = "SELECT FechaTransaccion, NumeroMovimiento AS NumeroSolicitud, DescripcionMovimiento, SubTotal, MontoIva, MontoRetenciones, MontoSolicitud, Anticipo From IndiceSolicitudPago "
    Me.DtaIndice.Refresh
 ElseIf Me.OptActivos.Value = True Then
    Me.DtaIndice.ConnectionString = Conexion
    Me.DtaIndice.RecordSource = "SELECT FechaTransaccion, NumeroMovimiento AS NumeroSolicitud, DescripcionMovimiento, SubTotal, MontoIva, MontoRetenciones, MontoSolicitud, Anticipo From IndiceSolicitudPago Where (Activo = 1) And (Procesado = 0) And (Anulado = 0)"
    Me.DtaIndice.Refresh
 ElseIf Me.OptAnulados.Value = True Then
    Me.DtaIndice.ConnectionString = Conexion
    Me.DtaIndice.RecordSource = "SELECT FechaTransaccion, NumeroMovimiento AS NumeroSolicitud, DescripcionMovimiento, SubTotal, MontoIva, MontoRetenciones, MontoSolicitud, Anticipo From IndiceSolicitudPago Where (Anulado = 1)"
    Me.DtaIndice.Refresh
 ElseIf Me.OptProcesados.Value = True Then
    Me.DtaIndice.ConnectionString = Conexion
    Me.DtaIndice.RecordSource = "SELECT FechaTransaccion, NumeroMovimiento AS NumeroSolicitud, DescripcionMovimiento, SubTotal, MontoIva, MontoRetenciones, MontoSolicitud, Anticipo From IndiceSolicitudPago Where (Procesado = 1)"
    Me.DtaIndice.Refresh
 
 End If


End Sub

Private Sub OptActivos_Click()
ActualizaGrid
End Sub

Private Sub OptAnulados_Click()
ActualizaGrid
End Sub

Private Sub OptProcesados_Click()
ActualizaGrid
End Sub

Private Sub OptTodos_Click()
ActualizaGrid
End Sub
