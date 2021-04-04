VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmTasa2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tasa"
   ClientHeight    =   4710
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   5280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   5280
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   4230
      Left            =   0
      Picture         =   "frmTasa2.frx":0000
      ScaleHeight     =   4170
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   0
      Width           =   1275
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   4320
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Height          =   375
      Left            =   480
      Top             =   6600
      Width           =   4215
      _ExtentX        =   7435
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
      Caption         =   "datPrimaryRS"
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
      Left            =   480
      Top             =   6960
      Width           =   4095
      _ExtentX        =   7223
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
   Begin VB.PictureBox Picture1 
      Height          =   4250
      Left            =   120
      ScaleHeight     =   4185
      ScaleWidth      =   5055
      TabIndex        =   0
      Top             =   0
      Width           =   5115
      Begin TrueOleDBGrid80.TDBGrid DBGrTasas 
         Bindings        =   "frmTasa2.frx":1B51
         Height          =   4215
         Left            =   1080
         TabIndex        =   1
         Top             =   0
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   7435
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Fecha Tasas"
         Columns(0).DataField=   "FechaTasas"
         Columns(0).NumberFormat=   "Short Date"
         Columns(0).EditMask=   "##/##/####"
         Columns(0).EditMaskUpdate=   -1  'True
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Monto Cordobas"
         Columns(1).DataField=   "MontoCordobas"
         Columns(1).NumberFormat=   "General Number"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).Caption=   "Tasas de Cambio"
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
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
         Appearance      =   2
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
         PictureStandardRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         PictureStandardRow(0)=   "bHQAAO4BAABCTe4BAAAAAAAANgAAACgAAAAOAAAACgAAAAEAGAAAAAAAuAEAAAAAAAAAAAAAAAAA"
         PictureStandardRow(1)=   "AAAAAADGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8YAAMbHxgAAAAAA"
         PictureStandardRow(2)=   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMbHxgAAxsfG////hIaEhIaEhIaEhIaEhIaE"
         PictureStandardRow(3)=   "hIaEhIaEhIaEhIaEhIaEAAAAxsfGAADGx8b////Gx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8aE"
         PictureStandardRow(4)=   "hoQAAADGx8YAAMbHxv///8bHxsbHxsbHxsbHxsbHxsbHxsbHxsbHxsbHxoSGhAAAAMbHxgAAxsfG"
         PictureStandardRow(5)=   "////xsfGxsfGxsfGxsfGxsfGxsfGxsfGxsfGxsfGhIaEAAAAxsfGAADGx8b////Gx8bGx8bGx8bG"
         PictureStandardRow(6)=   "x8bGx8bGx8bGx8bGx8bGx8aEhoQAAADGx8YAAMbHxv///8bHxsbHxsbHxsbHxsbHxsbHxsbHxsbH"
         PictureStandardRow(7)=   "xsbHxoSGhAAAAMbHxgAAxsfG////////////////////////////////////////////AAAAxsfG"
         PictureStandardRow(8)=   "AADGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8YAAA=="
         PictureStandardRow.vt=   9
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HBFD6DD&,.fgcolor=&H0&,.bold=-1"
         _StyleDefs(22)  =   ":id=22,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(23)  =   ":id=22,.fontname=MS Sans Serif"
         _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bgcolor=&HBFD6DD&"
         _StyleDefs(25)  =   ":id=14,.fgcolor=&H0&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(26)  =   ":id=14,.strikethrough=0,.charset=0"
         _StyleDefs(27)  =   ":id=14,.fontname=MS Sans Serif"
         _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(45)  =   "Named:id=33:Normal"
         _StyleDefs(46)  =   ":id=33,.parent=0"
         _StyleDefs(47)  =   "Named:id=34:Heading"
         _StyleDefs(48)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(49)  =   ":id=34,.wraptext=-1"
         _StyleDefs(50)  =   "Named:id=35:Footing"
         _StyleDefs(51)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(52)  =   "Named:id=36:Selected"
         _StyleDefs(53)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(54)  =   "Named:id=37:Caption"
         _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(56)  =   "Named:id=38:HighlightRow"
         _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(58)  =   "Named:id=39:EvenRow"
         _StyleDefs(59)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(60)  =   "Named:id=40:OddRow"
         _StyleDefs(61)  =   ":id=40,.parent=33"
         _StyleDefs(62)  =   "Named:id=41:RecordSelector"
         _StyleDefs(63)  =   ":id=41,.parent=34"
         _StyleDefs(64)  =   "Named:id=42:FilterBar"
         _StyleDefs(65)  =   ":id=42,.parent=33"
      End
      Begin VB.Image Image1 
         Height          =   4170
         Left            =   0
         Picture         =   "frmTasa2.frx":1B6C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmTasa2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdPegar_Click()
FrmCambioMoneda.MaskEdFecha.Text = Format(datPrimaryRS.Recordset("FechaDia"), "dd/mm/yyyy")
FrmCambioMoneda.MaskEdMonto.Text = Format(datPrimaryRS.Recordset("MontoDia"), "dd/mm/yyyy")
Unload Me
End Sub

Private Sub DBGrTasas_BeforeUpdate(Cancel As Integer)
  Me.DBGrTasas.Columns(1).Text = Format(CDbl(Me.DBGrTasas.Columns(1).Text), "##,##0.0000")
End Sub

Private Sub Form_Activate()
On Error GoTo TipoErrs

If Not CodigoUsuario = 0 Then
 Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tasa Cambio'))"
 Me.DtaNacceso.Refresh
 If Me.DtaNacceso.Recordset.EOF Then
   Me.cmdAdd.Enabled = False
   Me.DBGrTasas.Columns(0).Locked = True
   Me.DBGrTasas.Columns(1).Locked = True
  ' Me.DBGrTasas.Columns(2).Locked = True
 End If
 Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tasa Cambio'))"
 Me.DtaNacceso.Refresh
 If Me.DtaNacceso.Recordset.EOF Then
   Me.cmdDelete.Enabled = False
   'Me.DBGrTasas.Columns(0).Locked = True
   'Me.DBGrTasas.Columns(1).Locked = True
   'Me.DBGrTasas.Columns(2).Locked = True
 End If
End If
Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub Form_Load()
On Error GoTo TipoErrs
MDIPrimero.Skin1.ApplySkin hWnd
' Me.Picture2.Picture = Me.Image1.Picture
 Me.DBGrTasas.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DBGrTasas.OddRowStyle.BackColor = &H80000005
 Me.DBGrTasas.AlternatingRowStyle = True


With Me.DtaNacceso
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Cuentas"
End With

With Me.datPrimaryRS
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Tasas"
End With

Me.top = 1000
Me.Left = 4000
DBGrTasas.Columns(0).Locked = False
'Me.datPrimaryRS.DatabaseName = Ruta
Me.datPrimaryRS.ConnectionString = Conexion
Me.datPrimaryRS.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas ORDER BY Tasas.FechaTasas "
Me.datPrimaryRS.Refresh

Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo TipoErrs



'Dim Encontrado As BookmarkEnum
Dim Fecha As String
Dim NumFecha As Long
If Tasa = False Then
  Fecha = FrmTransacciones.TxtFecha.Value
Else
 Fecha = Format(Now, "dd/mm/yyyy")
End If
Fecha = Format(Fecha, "yyyy/mm/dd")
datPrimaryRS.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE(FechaTasas = CONVERT(DATETIME, '" & Fecha & "', 102)) ORDER BY FechaTasas"
'datPrimaryRS.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) = '" & NumFecha & "'))ORDER BY Tasas.FechaTasas"
datPrimaryRS.Refresh

If Not datPrimaryRS.Recordset.EOF Then
   Fecha = Format(datPrimaryRS.Recordset("FechaTasas"), "dd/mm/yyyy")
   
    Encontrado = True
    Cambio = datPrimaryRS.Recordset("MontoCordobas")
   ' MDIPrimero.StatusBar2.Panels(2) = "Tasa Cordobas: " & Format(Cambio, "##,##0.00") & "    " & "Tasa Libras: " & Format(datPrimaryRS.Recordset("MontoLibras"), "##,##0.00")
    MDIPrimero.StatusBar2.Panels(2) = "Tasa Cordobas: " & Format(Cambio, "##,##0.0000")
End If
 
 If Not Encontrado Then
   MsgBox "La Tasa de Hoy no ha sido grabada"
   Cancel = 100
 End If
Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'Esto cambiará el tamaño de la cuadrícula al cambiar el tamaño del formulario
  grdDataGrid.Height = Me.ScaleHeight - datPrimaryRS.Height - 30 - picButtons.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub


Private Sub cmdAdd_Click()
'  On Error GoTo AddErr


  Me.datPrimaryRS.Recordset.AddNew
'  Me.datPrimaryRS.Recordset.MoveLast
  Me.DBGrTasas.SetFocus
 If Tasa = False Then
   Me.DBGrTasas.Columns(0).Text = FrmTransacciones.TxtFecha.Value
   Me.DBGrTasas.Columns(1).Text = 0
 Else
  Me.DBGrTasas.Columns(0).Text = Format(Now, "dd/mm/yyyy")
  Me.DBGrTasas.Columns(1).Text = 0
'  Me.DBGrTasas.Columns(2).Text = 0
 End If
  Exit Sub
AddErr:
  MsgBox err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  Exit Sub
RefreshErr:
  MsgBox err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  datPrimaryRS.Refresh
  Exit Sub
UpdateErr:
  MsgBox err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

