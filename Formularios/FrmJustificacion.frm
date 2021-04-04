VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form FrmJustificacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos de la Justificacion"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   Icon            =   "FrmJustificacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   9360
      TabIndex        =   20
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton CmdSiguiente 
      Caption         =   "Siguiente"
      Height          =   495
      Left            =   5160
      TabIndex        =   19
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "Anterior"
      Height          =   495
      Left            =   3840
      TabIndex        =   18
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Grabar"
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   5040
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel Label1 
      Height          =   375
      Left            =   840
      OleObjectBlob   =   "FrmJustificacion.frx":628A
      TabIndex        =   9
      Top             =   240
      Width           =   8775
   End
   Begin MSAdodcLib.Adodc DtaContratista 
      Height          =   375
      Left            =   7080
      Top             =   7080
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
   Begin MSAdodcLib.Adodc DtaMovimiento 
      Height          =   330
      Left            =   4800
      Top             =   7080
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "DtaMovimiento"
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
      Height          =   330
      Left            =   1920
      Top             =   7080
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
   Begin VB.Frame Frame3 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   10455
      Begin MSMask.MaskEdBox TxtFechaIni 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtFechaT 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtFechaC 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtCosto 
         Height          =   285
         Left            =   4440
         TabIndex        =   5
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtCobro 
         Height          =   285
         Left            =   7680
         TabIndex        =   6
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPagado 
         Height          =   285
         Left            =   4560
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPor 
         Height          =   285
         Left            =   9000
         TabIndex        =   8
         Top             =   1080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "0%"
         PromptChar      =   "_"
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmJustificacion.frx":6314
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmJustificacion.frx":638C
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmJustificacion.frx":6404
         TabIndex        =   12
         Top             =   1200
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "FrmJustificacion.frx":6486
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "FrmJustificacion.frx":6500
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   9000
         OleObjectBlob   =   "FrmJustificacion.frx":6570
         TabIndex        =   15
         Top             =   840
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   7680
         OleObjectBlob   =   "FrmJustificacion.frx":65E4
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
   End
   Begin TrueDBGrid80.TDBGrid DBGMovimiento 
      Bindings        =   "FrmJustificacion.frx":6666
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   3625
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
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   10862530
      RowDividerColor =   10862530
      RowSubDividerColor=   10862530
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
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
End
Attribute VB_Name = "FrmJustificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAnterior_Click()
On Error GoTo TipoErrs
Dim Respuesta As Integer
FrmContactos.DtaContratista.Recordset.MovePrevious
If FrmContactos.DtaContratista.Recordset.BOF Then
   FrmContactos.DtaContratista.Recordset.MoveNext
   MsgBox "Este es el Primer Registro", vbInfoContabilidadtion, "Control de Cuentas Contabilidad"
Else
 FrmContactos.DBContratista.Text = FrmContactos.DtaContratista.Recordset("CodigoCuenta")
 If Not IsNull(FrmContactos.DtaContratista.Recordset("MontoAcordado")) Then
 MontoAcordado = FrmContactos.DtaContratista.Recordset("MontoAcordado")
 FrmJustificacion.TxtCosto = FrmContactos.DtaContratista.Recordset("MontoAcordado")
 Else
  MontoAcordado = 0
 FrmJustificacion.TxtCosto = "0"
End If
FrmJustificacion.TxtFechaIni = FrmContactos.TxtFechaContrata
FrmJustificacion.TxtFechaT = FrmContactos.TxtFechaFinaliza
FrmJustificacion.Label1 = FrmContactos.TxtNombre.Text
 
 MontoCobrado = 0
Me.DtaConsulta.RecordSource = "SELECT AdelantosJustifica.CodCuenta, AdelantosJustifica.MontoAnticipo From AdelantosJustifica Where (((AdelantosJustifica.CodCuenta) = '" & FrmContactos.DBContratista.Text & "'))"
Me.DtaConsulta.Refresh
Do While Not Me.DtaConsulta.Recordset.EOF
 MontoCobrado = Me.DtaConsulta.Recordset!MontoAnticipo + MontoCobrado

 Me.DtaConsulta.Recordset.MoveNext
Loop

 Me.TxtCobro.Text = MontoCobrado


MontoPendiente = MontoAcordado - MontoCobrado
Me.TxtPagado.Text = MontoPendiente
If Not MontoAcordado = 0 Then
 Me.TxtPor.Text = MontoPendiente / MontoAcordado
Else
 Me.TxtPor.Text = 0
End If

Me.DtaMovimiento.RecordSource = "SELECT AdelantosJustifica.FechaAnticipo, AdelantosJustifica.NTransaccion, AdelantosJustifica.RefVoucher, AdelantosJustifica.CodigoCuenta, AdelantosJustifica.DateClearerance, AdelantosJustifica.MisCode, AdelantosJustifica.CodCuenta, AdelantosJustifica.MontoAnticipo From AdelantosJustifica Where (((AdelantosJustifica.CodCuenta) = '" & FrmContactos.DBContratista.Text & "')) ORDER BY AdelantosJustifica.FechaAnticipo"

'Me.DtaMovimiento.RecordSource = "SELECT AdelantosJustifica.FechaAnticipo, AdelantosJustifica.NTransaccion, AdelantosJustifica.RefVoucher, AdelantosJustifica.MontoAnticipo, AdelantosJustifica.CodCuenta From AdelantosJustifica WHERE (((AdelantosJustifica.CodCuenta)='" & FrmContactos.DBContratista.Text & "'))"
Me.DtaMovimiento.Refresh
If Not DtaMovimiento.Recordset.EOF Then
 Me.DtaMovimiento.Recordset.MoveLast
End If
If Not (Me.DtaMovimiento.Recordset.EOF) Then
 Me.TxtFechaC.Text = Me.DtaMovimiento.Recordset!FechaAnticipo
Else
 Me.TxtFechaC.Text = "  /  /    "
End If
Me.Label1.Caption = FrmContactos.TxtNombre
Me.DBGMovimiento.Columns(0).Width = "1500"
Me.DBGMovimiento.Columns(1).Width = "1200"
Me.DBGMovimiento.Columns(2).Width = "1200"
Me.DBGMovimiento.Columns(3).Width = "1500"
Me.DBGMovimiento.Columns(3).Button = True
Me.DBGMovimiento.Columns(4).Button = True
Me.DBGMovimiento.Columns(5).Button = True
Me.DBGMovimiento.Columns(3).Caption = "Suspense Code"
Me.DBGMovimiento.Columns(6).Visible = False
Me.DBGMovimiento.Columns(7).NumberFormat = "##,##0.00"
Me.DBGMovimiento.Columns(7).Locked = True


End If
   Exit Sub
TipoErrs:
   ControlErrores
End Sub

Private Sub CmdGrabar_Click()
Criterio = "CodigoCuenta='" & FrmContactos.DBContratista.Text & "'"
Me.DtaContratista.Recordset.Find (Criterio)
If Not DtaContratista.Recordset.EOF Then
  'Me.DtaContratista.Recordset.Edit
   Me.DtaContratista.Recordset("MontoAcordado") = Me.TxtCosto.Text
  Me.DtaContratista.Recordset.Update

End If
 Me.DtaMovimiento.RecordSource = "SELECT AdelantosJustifica.FechaAnticipo, AdelantosJustifica.NTransaccion, AdelantosJustifica.RefVoucher, AdelantosJustifica.CodigoCuenta, AdelantosJustifica.DateClearerance, AdelantosJustifica.MisCode, AdelantosJustifica.CodCuenta, AdelantosJustifica.MontoAnticipo From AdelantosJustifica Where (((AdelantosJustifica.CodCuenta) = '" & FrmContactos.DBContratista.Text & "')) ORDER BY AdelantosJustifica.FechaAnticipo"
 ' Me.DtaMovimiento.RecordSource = "SELECT AdelantosJustifica.FechaAnticipo, AdelantosJustifica.NTransaccion, AdelantosJustifica.RefVoucher, AdelantosJustifica.MontoAnticipo, AdelantosJustifica.CodCuenta From AdelantosJustifica WHERE (((AdelantosJustifica.CodCuenta)='*'))"
  Me.DtaMovimiento.Refresh
  
Me.DBGMovimiento.Columns(0).Width = "1500"
Me.DBGMovimiento.Columns(1).Width = "1200"
Me.DBGMovimiento.Columns(2).Width = "1200"
Me.DBGMovimiento.Columns(3).Width = "1500"
Me.DBGMovimiento.Columns(3).Button = True
Me.DBGMovimiento.Columns(4).Button = True
Me.DBGMovimiento.Columns(5).Button = True
Me.DBGMovimiento.Columns(3).Caption = "Suspense Code"
Me.DBGMovimiento.Columns(7).NumberFormat = "##,##0.00"
Me.DBGMovimiento.Columns(7).Locked = True
Me.DBGMovimiento.Columns(6).Visible = False
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSiguiente_Click()
On Error GoTo TipoErrs
Dim Respuesta As Integer
FrmContactos.DtaContratista.Recordset.MoveNext
If FrmContactos.DtaContratista.Recordset.EOF Then
   FrmContactos.DtaContratista.Recordset.MovePrevious
   MsgBox "Este es el Ultimo Registro", vbInfoContabilidadtion, "Control de contratista Contabilidad"
Else
 FrmContactos.DBContratista.Text = FrmContactos.DtaContratista.Recordset("CodigoCuenta")
 If Not IsNull(FrmContactos.DtaContratista.Recordset("MontoAcordado")) Then
 MontoAcordado = FrmContactos.DtaContratista.Recordset("MontoAcordado")
 FrmJustificacion.TxtCosto = FrmContactos.DtaContratista.Recordset("MontoAcordado")
 Else
  MontoAcordado = 0
  FrmJustificacion.TxtCosto = "0"
 End If
FrmJustificacion.TxtFechaIni = FrmContactos.TxtFechaContrata
FrmJustificacion.TxtFechaT = FrmContactos.TxtFechaFinaliza
FrmJustificacion.Label1 = FrmContactos.TxtNombre.Text

 MontoCobrado = 0
Me.DtaConsulta.RecordSource = "SELECT AdelantosJustifica.CodCuenta, AdelantosJustifica.MontoAnticipo From AdelantosJustifica Where (((AdelantosJustifica.CodCuenta) = '" & FrmContactos.DBContratista.Text & "'))"
Me.DtaConsulta.Refresh
Do While Not Me.DtaConsulta.Recordset.EOF
 MontoCobrado = Me.DtaConsulta.Recordset!MontoAnticipo + MontoCobrado

 Me.DtaConsulta.Recordset.MoveNext
Loop

 Me.TxtCobro.Text = MontoCobrado
 
MontoPendiente = MontoAcordado - MontoCobrado
Me.TxtPagado.Text = MontoPendiente
If Not MontoAcordado = 0 Then
 Me.TxtPor.Text = MontoPendiente / MontoAcordado
Else
 Me.TxtPor.Text = 0
End If


Me.DtaMovimiento.RecordSource = "SELECT AdelantosJustifica.FechaAnticipo, AdelantosJustifica.NTransaccion, AdelantosJustifica.RefVoucher, AdelantosJustifica.CodigoCuenta, AdelantosJustifica.DateClearerance, AdelantosJustifica.MisCode, AdelantosJustifica.CodCuenta, AdelantosJustifica.MontoAnticipo From AdelantosJustifica Where (((AdelantosJustifica.CodCuenta) = '" & FrmContactos.DBContratista.Text & "')) ORDER BY AdelantosJustifica.FechaAnticipo"
'Me.DtaMovimiento.RecordSource = "SELECT AdelantosJustifica.FechaAnticipo, AdelantosJustifica.NTransaccion, AdelantosJustifica.RefVoucher, AdelantosJustifica.MontoAnticipo, AdelantosJustifica.CodCuenta From AdelantosJustifica WHERE (((AdelantosJustifica.CodCuenta)='" & FrmContactos.DBContratista.Text & "'))"
Me.DtaMovimiento.Refresh
If Not DtaMovimiento.Recordset.EOF Then
 Me.DtaMovimiento.Recordset.MoveLast
End If
If Not (Me.DtaMovimiento.Recordset.EOF) Then
 Me.TxtFechaC.Text = Me.DtaMovimiento.Recordset!FechaAnticipo
Else
 Me.TxtFechaC.Text = "  /  /    "
End If
Me.Label1.Caption = FrmContactos.TxtNombre
Me.DBGMovimiento.Columns(0).Width = "1500"
Me.DBGMovimiento.Columns(1).Width = "1200"
Me.DBGMovimiento.Columns(2).Width = "1200"
Me.DBGMovimiento.Columns(3).Width = "1500"
Me.DBGMovimiento.Columns(3).Button = True
Me.DBGMovimiento.Columns(4).Button = True
Me.DBGMovimiento.Columns(5).Button = True
Me.DBGMovimiento.Columns(3).Caption = "Suspense Code"
Me.DBGMovimiento.Columns(6).Visible = False
Me.DBGMovimiento.Columns(7).NumberFormat = "##,##0.00"
Me.DBGMovimiento.Columns(7).Locked = True



End If
   Exit Sub
TipoErrs:
   ControlErrores
End Sub

Private Sub DBGMovimiento_ButtonClick(ByVal ColIndex As Integer)
 Select Case ColIndex
   Case 3
    QueProducto = "ContratistasM"
    FrmConsulta.Show 1
   Case 4
    FrmMes.Show 1
   Case 5
    QueProducto = "MisCode"
    FrmConsulta.Show 1
    
 End Select
End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd
Dim MontoCobrado As Double, MontoPendiente As Double
With Me.DtaContratista
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from Contactos"
   .Refresh
End With

With Me.DtaMovimiento
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaConsulta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With



Me.DtaConsulta.RecordSource = "SELECT AdelantosJustifica.CodCuenta, AdelantosJustifica.MontoAnticipo From AdelantosJustifica Where (((AdelantosJustifica.CodCuenta) = '" & FrmContactos.DBContratista.Text & "'))"
Me.DtaConsulta.Refresh
Do While Not Me.DtaConsulta.Recordset.EOF
 MontoCobrado = Me.DtaConsulta.Recordset!MontoAnticipo + MontoCobrado

 Me.DtaConsulta.Recordset.MoveNext
Loop

 Me.TxtCobro.Text = MontoCobrado
 

MontoPendiente = MontoAcordado - MontoCobrado
Me.TxtPagado.Text = MontoPendiente
If Not MontoAcordado = 0 Then
 Me.TxtPor.Text = MontoPendiente / MontoAcordado
End If


Me.DtaMovimiento.RecordSource = "SELECT AdelantosJustifica.FechaAnticipo, AdelantosJustifica.NTransaccion, AdelantosJustifica.RefVoucher, AdelantosJustifica.CodigoCuenta, AdelantosJustifica.DateClearerance, AdelantosJustifica.MisCode, AdelantosJustifica.CodCuenta, AdelantosJustifica.MontoAnticipo From AdelantosJustifica Where (((AdelantosJustifica.CodCuenta) = '" & FrmContactos.DBContratista.Text & "')) ORDER BY AdelantosJustifica.FechaAnticipo"
'Me.DtaMovimiento.RecordSource = "SELECT AdelantosJustifica.FechaAnticipo, AdelantosJustifica.NTransaccion, AdelantosJustifica.RefVoucher, AdelantosJustifica.MontoAnticipo, AdelantosJustifica.CodCuenta From AdelantosJustifica WHERE (((AdelantosJustifica.CodCuenta)='" & FrmContactos.DBContratista.Text & "'))"
Me.DtaMovimiento.Refresh
If Not DtaMovimiento.Recordset.EOF Then
 Me.DtaMovimiento.Recordset.MoveLast
End If
If Not (Me.DtaMovimiento.Recordset.EOF) Then
 Me.TxtFechaC.Text = Me.DtaMovimiento.Recordset!FechaAnticipo
End If
Me.DBGMovimiento.Columns(0).Width = "1500"
Me.DBGMovimiento.Columns(1).Width = "1200"
'Me.DBGMovimiento.Columns(2).Width = "1200"
'Me.DBGMovimiento.Columns(3).Width = "1500"
'Me.DBGMovimiento.Columns(3).Button = True
'Me.DBGMovimiento.Columns(4).Button = True
'Me.DBGMovimiento.Columns(5).Button = True
'Me.DBGMovimiento.Columns(3).Caption = "Suspense Code"
'Me.DBGMovimiento.Columns(7).NumberFormat = "##,##0.00"
'Me.DBGMovimiento.Columns(7).Locked = True
'Me.DBGMovimiento.Columns(6).Visible = False
End Sub

Private Sub TxtCosto_Change()
If IsNumeric(Me.TxtCosto.Text) Then
 MontoAcordado = Me.TxtCosto.Text
End If
Me.DtaConsulta.RecordSource = "SELECT AdelantosJustifica.CodCuenta, AdelantosJustifica.MontoAnticipo From AdelantosJustifica Where (((AdelantosJustifica.CodCuenta) = '" & FrmContactos.DBContratista.Text & "'))"
Me.DtaConsulta.Refresh
Do While Not Me.DtaConsulta.Recordset.EOF
 MontoCobrado = Me.DtaConsulta.Recordset!MontoAnticipo + MontoCobrado

 Me.DtaConsulta.Recordset.MoveNext
Loop

 Me.TxtCobro.Text = MontoCobrado
MontoPendiente = MontoAcordado - MontoCobrado
Me.TxtPagado.Text = MontoPendiente
If Not MontoAcordado = 0 Then
 Me.TxtPor.Text = MontoPendiente / MontoAcordado
Else
 Me.TxtPor.Text = 0
End If


Me.DtaMovimiento.RecordSource = "SELECT AdelantosJustifica.FechaAnticipo, AdelantosJustifica.NTransaccion, AdelantosJustifica.RefVoucher, AdelantosJustifica.CodigoCuenta, AdelantosJustifica.DateClearerance, AdelantosJustifica.MisCode, AdelantosJustifica.CodCuenta, AdelantosJustifica.MontoAnticipo From AdelantosJustifica Where (((AdelantosJustifica.CodCuenta) = '" & FrmContactos.DBContratista.Text & "')) ORDER BY AdelantosJustifica.FechaAnticipo"
'Me.DtaMovimiento.RecordSource = "SELECT AdelantosJustifica.FechaAnticipo, AdelantosJustifica.NTransaccion, AdelantosJustifica.RefVoucher, AdelantosJustifica.MontoAnticipo, AdelantosJustifica.CodCuenta From AdelantosJustifica WHERE (((AdelantosJustifica.CodCuenta)='" & FrmContactos.DBContratista.Text & "'))"
Me.DtaMovimiento.Refresh
If Not DtaMovimiento.Recordset.EOF Then
 Me.DtaMovimiento.Recordset.MoveLast
End If
If Not (Me.DtaMovimiento.Recordset.EOF) Then
 Me.TxtFechaC.Text = Me.DtaMovimiento.Recordset!FechaAnticipo
Else
 Me.TxtFechaC.Text = "  /  /    "
End If
'Me.Label1.Caption = FrmContactos.TxtNombre
'Me.DBGMovimiento.Columns(0).Width = "1500"
'Me.DBGMovimiento.Columns(1).Width = "1200"
'Me.DBGMovimiento.Columns(2).Width = "1200"
'Me.DBGMovimiento.Columns(3).Width = "1500"
'Me.DBGMovimiento.Columns(3).Button = True
'Me.DBGMovimiento.Columns(5).Button = True
'Me.DBGMovimiento.Columns(4).Button = True
'Me.DBGMovimiento.Columns(3).Caption = "Suspense Code"
'Me.DBGMovimiento.Columns(6).Visible = False
'Me.DBGMovimiento.Columns(7).NumberFormat = "##,##0.00"
'Me.DBGMovimiento.Columns(7).Locked = True
End Sub
