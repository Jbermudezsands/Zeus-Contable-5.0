VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FrmTransaccionCob 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro Cuentas x Cobrar , Cuentas x Pagar"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoBusca 
      Height          =   375
      Left            =   2160
      Top             =   4680
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
      Caption         =   "AdoBusca"
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
      Left            =   2160
      Top             =   4080
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
   Begin VB.CommandButton CmdPegar 
      Caption         =   "Pegar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   3360
      Width           =   1455
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGridConsulta 
      Bindings        =   "TransaccionCtasCob.frx":0000
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5530
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "No Factura"
      Columns(0).DataField=   "FacturaNo"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripcion Movimiento"
      Columns(1).DataField=   "DescripcionMovimiento"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Fecha Vencimiento"
      Columns(2).DataField=   "FechaVence"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Saldo"
      Columns(3).DataField=   "Saldo"
      Columns(3).NumberFormat=   "Fixed"
      Columns(3).EditMask=   "##,##0.00"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   1085
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).DividerColor=   16315377
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=138,.bold=0,.fontsize=825,.italic=0"
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
End
Attribute VB_Name = "FrmTransaccionCob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdPegar_Click()
FrmTransacciones.DBGTransacciones.Columns(4).Text = Me.TDBGridConsulta.Columns(0).Text
Unload Me
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim SQL As String, CodigoCuenta As String, TipoCuenta As String
MDIPrimero.Skin1.ApplySkin hWnd
 Me.TDBGridConsulta.EvenRowStyle.BackColor = &H80FFFF
 Me.TDBGridConsulta.OddRowStyle.BackColor = &HC0FFFF
 Me.TDBGridConsulta.AlternatingRowStyle = True

With Me.AdoConsulta
   .ConnectionString = Conexion
End With

With Me.AdoBusca
   .ConnectionString = Conexion
End With

CodigoCuenta = FrmTransacciones.DBGTransacciones.Columns(0).Text
Me.AdoBusca.RecordSource = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo From Cuentas WHERE (CodCuentas = '" & CodigoCuenta & "')"
Me.AdoBusca.Refresh
If Not Me.AdoBusca.Recordset.EOF Then
   TipoCuenta = Me.AdoBusca.Recordset("TipoCuenta")

End If

                      
                      

If TipoCuenta = "Cuentas x Cobrar" Then
   SQL = "SELECT Transacciones.CodCuentas AS CodCuentas, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, " & _
      "MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.FechaVence) AS FechaVence, " & _
      "MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Cuentas.TipoCuenta) AS TipoCuenta, Transacciones.FacturaNo AS FacturaNo, " & _
      "SUM(Transacciones.Debito * Transacciones.TCambio) AS Debito, SUM(Transacciones.Credito * Transacciones.TCambio) AS Credito, " & _
      "SUM(Transacciones.Debito * Transacciones.TCambio - Transacciones.Credito * Transacciones.TCambio) As Saldo " & _
      "FROM         Transacciones INNER JOIN " & _
      "Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas " & _
      "GROUP BY Transacciones.FacturaNo, Transacciones.CodCuentas " & _
      "HAVING      (MAX(Transacciones.FechaVence) <> CONVERT(DATETIME, '1900-01-01 00:00:00', 102)) AND (MAX(Cuentas.TipoCuenta) = '" & TipoCuenta & "') AND (Transacciones.CodCuentas = '" & CodigoCuenta & "') " & _
      "ORDER BY Transacciones.FacturaNo "
Else
  SQL = "SELECT Transacciones.CodCuentas AS CodCuentas, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, " & _
      "MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.FechaVence) AS FechaVence, " & _
      "MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Cuentas.TipoCuenta) AS TipoCuenta, Transacciones.FacturaNo AS FacturaNo, " & _
      "SUM(Transacciones.Debito * Transacciones.TCambio) AS Debito, SUM(Transacciones.Credito * Transacciones.TCambio) AS Credito, " & _
      "SUM(Transacciones.Debito * Transacciones.TCambio - Transacciones.Credito * Transacciones.TCambio) As Saldo " & _
      "FROM         Transacciones INNER JOIN " & _
      "Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas " & _
      "GROUP BY Transacciones.FacturaNo, Transacciones.CodCuentas " & _
      "HAVING      (MAX(Transacciones.FechaVence) <> CONVERT(DATETIME, '1900-01-01 00:00:00', 102)) AND (MAX(Cuentas.TipoCuenta) = '" & TipoCuenta & "') AND (Transacciones.CodCuentas = '" & CodigoCuenta & "') " & _
      "ORDER BY Transacciones.FacturaNo "


End If
Me.AdoConsulta.RecordSource = SQL
Me.AdoConsulta.Refresh


End Sub
