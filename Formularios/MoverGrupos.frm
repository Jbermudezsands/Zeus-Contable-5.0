VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmMoverGrupos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Moviendo Grupo de Cuentas"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ProgressBar osProgress1 
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   5520
      Width           =   7095
      _Version        =   786432
      _ExtentX        =   12515
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin VB.CommandButton CmdCambiar 
      Cancel          =   -1  'True
      Caption         =   "Cambiar"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9000
      TabIndex        =   9
      Top             =   5520
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc AdoCuentas 
      Height          =   375
      Left            =   720
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
   Begin TrueOleDBGrid80.TDBGrid TDBGridMover 
      Bindings        =   "MoverGrupos.frx":0000
      Height          =   3495
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6165
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Codigo Cuenta"
      Columns(0).DataField=   "CodCuentas"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripcion Cuenta"
      Columns(1).DataField=   "DescripcionCuentas"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Tipo Cuenta"
      Columns(2).DataField=   "TipoCuenta"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Grupo de Cuentas"
      Columns(3).DataField=   "DescripcionGrupo"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   1085
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).DividerColor=   16315377
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3519"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3440"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=1"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=4419"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4339"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=5292"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=5212"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
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
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
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
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      Begin VB.TextBox TxtKeyGrupo 
         Height          =   285
         Left            =   4440
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton CmdConsultar 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   8400
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox TxtDescripcionGrupo 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   8415
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
         Left            =   9600
         Picture         =   "MoverGrupos.frx":0019
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox TxtCodigoInicio 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   480
         OleObjectBlob   =   "MoverGrupos.frx":0167
         TabIndex        =   5
         Top             =   960
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   495
         Left            =   120
         OleObjectBlob   =   "MoverGrupos.frx":01DB
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   4440
         TabIndex        =   2
         Top             =   360
         Width           =   4695
      End
   End
   Begin MSAdodcLib.Adodc AdoCambiar 
      Height          =   375
      Left            =   720
      Top             =   7080
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
      Caption         =   "AdoCambiar"
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
Attribute VB_Name = "FrmMoverGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub CmdBuscarEmpleado_Click()
QUIEN = "MoverGrupos"
FrmGrupoLista.Show 1
End Sub

Private Sub CmdCambiar_Click()
On Error GoTo TipoErrs
Dim i As Double, Cantidad As Double, Codigo As String
Dim KeyTipo As String, KeyGrupoCuenta As String, TipoCuenta As String


If Me.TxtCodigoInicio.Text = "" Then
  MsgBox "Debe Consultar el Grupo de Cuentas", vbCritical, "Sistema Contable"
  Exit Sub
End If

If Me.TxtDescripcionGrupo.Text = "" Then
  MsgBox "Necesita Selecionar el Grupo", vbCritical, "Sistema Contable"
  Exit Sub
End If

If Me.TxtKeyGrupo.Text = "" Then
  MsgBox "Necesita Selecionar el Grupo", vbCritical, "Sistema Contable"
  Exit Sub
End If


Codigo = Me.TxtCodigoInicio.Text

SQL = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo " & _
"From Cuentas WHERE (CodCuentas LIKE '" & Codigo & "%') ORDER BY CodCuentas"
AdoCambiar.RecordSource = SQL
AdoCambiar.Refresh
Me.AdoCambiar.Recordset.MoveLast
Cantidad = Me.AdoCambiar.Recordset.RecordCount
Me.AdoCambiar.Recordset.MoveFirst

Me.osProgress1.Visible = True

With Me.osProgress1
 .Min = 0
 .Max = Cantidad
 .Value = 0
 i = 1
  Do While Not AdoCambiar.Recordset.EOF
  
    KeyTipo = Mid(Me.TxtKeyGrupo.Text, 1, 1)
    KeyGrupoCuenta = Me.TxtKeyGrupo.Text
    TipoCuenta = Me.AdoCambiar.Recordset("Tipocuenta")
    If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
        TipoCuenta = "A"
    ElseIf TipoCuenta = "Otros Pasivos" Or TipoCuenta = "Cuentas x Pagar" Or TipoCuenta = "Pasivo" Then
        TipoCuenta = "B"
    ElseIf TipoCuenta = "Capital" Then
        TipoCuenta = "C"
    ElseIf TipoCuenta = "Costos" Then
        TipoCuenta = "G"
    ElseIf TipoCuenta = "Gastos" Then
        TipoCuenta = "O"
    ElseIf TipoCuenta = "Ingresos - Ventas" Then
        TipoCuenta = "D"
    ElseIf TipoCuenta = "Cuentas de Orden" Then
        TipoCuenta = "P"
    End If
   
 
    If KeyTipo = TipoCuenta Then
        Me.AdoCambiar.Recordset("DescripcionGrupo") = Me.TxtDescripcionGrupo.Text
        Me.AdoCambiar.Recordset("KeyGrupo") = Me.TxtKeyGrupo.Text
        Me.AdoCambiar.Recordset.Update
    End If
    
    Me.AdoCambiar.Recordset.MoveNext
    .Value = i
    i = i + 1
  Loop
End With

MsgBox "Cambio Correcto!!", vbExclamation, "Sistema Contable"
Me.osProgress1.Visible = False
SQL = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo " & _
"From Cuentas WHERE (CodCuentas LIKE '" & Codigo & "%') ORDER BY CodCuentas"
Me.AdoCuentas.RecordSource = SQL
Me.AdoCuentas.Refresh

Exit Sub
TipoErrs:
 MsgBox err.Description
End Sub

Private Sub CmdConsultar_Click()
Dim Codigo As String
Codigo = Me.TxtCodigoInicio.Text

SQL = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo " & _
"From Cuentas WHERE (CodCuentas LIKE '" & Codigo & "%') ORDER BY CodCuentas"
Me.AdoCuentas.RecordSource = SQL
Me.AdoCuentas.Refresh


End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd

 Me.TDBGridMover.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.TDBGridMover.OddRowStyle.BackColor = &H80000005
 Me.TDBGridMover.AlternatingRowStyle = True



With Me.AdoCuentas
   .ConnectionString = Conexion
End With

With Me.AdoCambiar
   .ConnectionString = Conexion
End With




End Sub

Private Sub Text1_Change()

End Sub

