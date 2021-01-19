VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FrmConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Cuentas"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10440
   FillColor       =   &H00C0FFFF&
   HelpContextID   =   6
   Icon            =   "FrmConsulCompra.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin TrueOleDBGrid80.TDBGrid DbgrProducto 
      Bindings        =   "FrmConsulCompra.frx":0E42
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4895
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
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=8,.bold=0,.fontsize=825,.italic=0"
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
   Begin MSAdodcLib.Adodc DtaProductos 
      Height          =   375
      Left            =   1200
      Top             =   5760
      Width           =   3975
      _ExtentX        =   7011
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
      Caption         =   "DtaProductos"
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
      Left            =   1200
      Top             =   5400
      Width           =   3975
      _ExtentX        =   7011
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
   Begin MSAdodcLib.Adodc DtaTasas 
      Height          =   375
      Left            =   1200
      Top             =   5040
      Width           =   3975
      _ExtentX        =   7011
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
      Caption         =   "DtaTasas"
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
      Caption         =   "&Pegar"
      Height          =   375
      Left            =   240
      MouseIcon       =   "FrmConsulCompra.frx":0E5D
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2280
      MouseIcon       =   "FrmConsulCompra.frx":129F
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton CmdOrden 
      Caption         =   "&Orden"
      Height          =   375
      Left            =   1200
      MouseIcon       =   "FrmConsulCompra.frx":16E1
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc AdoBusca 
      Height          =   375
      Left            =   1200
      Top             =   6480
      Width           =   3975
      _ExtentX        =   7011
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
End
Attribute VB_Name = "FrmConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnx As New ADODB.Connection
Private rs As New ADODB.Recordset, rsConexion As New ADODB.Recordset
Private Sql As String
Private modal As Boolean
Private getVal As Boolean
Private Id As Integer
Public Codigo As String, Cuenta As String, KeyPresupuesto As String


Private Sub DbgrProducto_PostEvent(ByVal MsgId As Integer)
 On Error GoTo TipoErrs
 
   If MsgId = 418 Then

'
'            If Me.DbgrProducto.Columns(0).FilterText = "" And Me.DbgrProducto.Columns(1).FilterText = "" And Me.DbgrProducto.Columns(2).FilterText = "" Then
'             rs.Filter = ""
'             Exit Sub
'            End If
        
        
            Dim col As TrueOleDBGrid80.Column
            Dim cols As TrueOleDBGrid80.Columns
            Dim Res As String
        
        
            On Error Resume Next
            Set cols = Me.DbgrProducto.Columns
            Dim c As Integer
        
        
            c = DbgrProducto.col
            DbgrProducto.HoldFields
            Sql = rs.Filter
            rs.Filter = getFilter(col, cols)
'            If rs.EOF Then
'              MsgBox "No Existen Registros", vbInformation, "Zeus Contabilidad"
'              Res = LimpiarFilter(col, cols)
'              rs.Filter = ""
'            End If

            DbgrProducto.col = c
            DbgrProducto.EditActive = True
  End If
  
  Exit Sub
TipoErrs:
 MsgBox err.Description
  

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set cnx = Nothing
Set rs = Nothing
Set rs2 = Nothing
Set rs3 = Nothing
End Sub


Private Sub CmdCancelar_Click()
On Error GoTo TipoErrs
Unload Me
Exit Sub
TipoErrs:
   ControlErrores
End Sub

Private Sub CmdOrden_Click()
On Error GoTo TipoErrs
Select Case QueProducto
      Case "Auxiliar"
         If Orden = True Then
         sqlconsulta = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.CodCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(0).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
         DbgrProducto.Columns(1).Width = 4200
            Orden = False
        Else
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
            Orden = True
         End If
         
      Case "CuentaDepreciacion"
         If Orden = True Then
         sqlconsulta = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.CodCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(0).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
         DbgrProducto.Columns(1).Width = 4200
            Orden = False
        Else
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
            Orden = True
         End If
         
      Case "CuentaGastos"
         If Orden = True Then
         sqlconsulta = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.CodCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(0).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
         DbgrProducto.Columns(1).Width = 4200
            Orden = False
        Else
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
            Orden = True
         End If
         
      Case "CuentaOriginal"
         If Orden = True Then
         sqlconsulta = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.CodCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(0).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
         DbgrProducto.Columns(1).Width = 4200
            Orden = False
        Else
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
            Orden = True
         End If
         

      Case "Periodo"
         If Orden = True Then
         sqlconsulta = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, , Cuentas.TipoMoneda, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE(Cuentas.TipoCuenta = 'Capital') ORDER BY Cuentas.CodCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(0).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
         DbgrProducto.Columns(1).Width = 4200
            Orden = False
        Else
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
            Orden = True
         End If
      
      Case "Cuenta"
         If Orden = True Then
         sqlconsulta = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.CodCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(0).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
         DbgrProducto.Columns(1).Width = 4200
            Orden = False
        Else
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
            Orden = True
         End If
         
      Case "CuentaReportes"
         If Orden = True Then
         sqlconsulta = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.CodCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(0).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
         DbgrProducto.Columns(1).Width = 4200
            Orden = False
        Else
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
            Orden = True
         End If
         
      Case "CuentaReportes2"
         If Orden = True Then
         sqlconsulta = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.CodCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(0).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
         DbgrProducto.Columns(1).Width = 4200
            Orden = False
        Else
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
            Orden = True
         End If
         
      Case "CuentaMayor"
         If Orden = True Then
         sqlconsulta = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.CodCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(0).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
         DbgrProducto.Columns(1).Width = 4200
            Orden = False
        Else
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
            Orden = True
         End If


      Case "Transferencia2"
         If Orden = True Then
         sqlconsulta = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.CodCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(0).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
         DbgrProducto.Columns(1).Width = 4200
            Orden = False
        Else
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
            Orden = True
         End If
      Case "Transferencia1"
         If Orden = True Then
         sqlconsulta = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.CodCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(0).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
         DbgrProducto.Columns(1).Width = 4200
            Orden = False
        Else
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
            Orden = True
         End If
         
      Case "ContratistasM"
         If Orden = True Then
         sqlconsulta = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.CodCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(0).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
         DbgrProducto.Columns(1).Width = 4200
            Orden = False
        Else
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
            Orden = True
         End If

        Case "MisCode"
         If Orden = True Then
         sqlconsulta = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.CodCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(0).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
         DbgrProducto.Columns(1).Width = 4200
            Orden = False
        Else
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
            Orden = True
         End If
   Case "AuxiliarTransacciones"
         If Orden = True Then
         sqlconsulta = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.CodCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(0).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
         DbgrProducto.Columns(1).Width = 4200
            Orden = False
        Else
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
            Orden = True
         End If

   Case "Transacciones"
         If Orden = True Then
         sqlconsulta = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.CodCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(0).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
         DbgrProducto.Columns(1).Width = 4200
            Orden = False
        Else
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
            Orden = True
         End If
   
      Case "Cheque"
         If Orden = True Then
         sqlconsulta = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.CodCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(0).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
         DbgrProducto.Columns(1).Width = 4200
            Orden = False
        Else
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
          Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
            Orden = True
         End If
         
   Case "ContratistaCheque"
         If Orden = True Then
         
         sqlconsulta = "SELECT Contactos.CodigoCuenta, Contactos.Beneficiario, Contactos.Ciudad, Contactos.Telefono From Contactos ORDER BY Contactos.Beneficiario"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(0).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 4200
         Me.DbgrProducto.Columns(3).Width = 2000
         DbgrProducto.Columns(0).Width = 2000
            Orden = False
        Else
         sqlconsulta = "SELECT Contactos.Beneficiario, Contactos.CodigoCuenta, Contactos.Ciudad, Contactos.Telefono From Contactos ORDER BY Contactos.CodigoCuenta"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 4200
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(1).Width = 2000
            Orden = True
         End If
      Case "Contratista"
         If Orden = True Then
         
         sqlconsulta = "SELECT Contactos.CodigoCuenta, Contactos.Beneficiario, Contactos.Ciudad, Contactos.Telefono From Contactos ORDER BY Contactos.Beneficiario"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(0).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 4200
         Me.DbgrProducto.Columns(3).Width = 2000
         DbgrProducto.Columns(0).Width = 2000
            Orden = False
        Else
         sqlconsulta = "SELECT Contactos.Beneficiario, Contactos.CodigoCuenta, Contactos.Ciudad, Contactos.Telefono From Contactos ORDER BY Contactos.CodigoCuenta"
         DtaProductos.RecordSource = sqlconsulta
         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 4200
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(1).Width = 2000
            Orden = True
         End If
         
    End Select
 Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub CmdPegar_Click()
'On Error GoTo TipoErrs
Dim valor1 As String, Valor As String, PrecioCosto As Integer
Dim Cantidad, Subtotal0, Subtotal1, Producto, Valores, Costo As Double
Dim CodigoP As String, Candena As String, Fecha As Long
Dim PrecioC As Double, CantidadC As Double, SubT As Double, Fechas1 As String, Fechas2 As String
Dim AnteriorSub As Double, cadena As String, Fechas As String, DescripcionMovimiento As String
Dim SqlDetalle As String, numero As String, TipoCuenta As String, ClaveMovimiento As String
Dim FechaIni As String, FechaFin As String, CodCuenta As String, NumeroProrrateo As Double
Dim mes As Double, Año As Double
  

  
  Codigo = ""
 'Busco el numero consecutivo de la Recepcion
  Select Case QueProducto
  
      Case "Presupuesto"
        If Not IsNull(rs("DescripcionGrupo")) Then
          If Not rs.EOF Then
           cadena = rs("DescripcionGrupo")
          End If
        End If
        
        KeyPresupuesto = rs("KeyGrupo")
        Codigo = CadenaNumeros(cadena)
  
      Case "CuentaContable"
       If Not IsNull(rs("CodCuentas")) Then
         If Not rs.EOF Then
          Cuenta = rs("CodCuentas")
          End If
       End If
  
      Case "Fuente"
       If Not IsNull(rs("Fuente")) Then
         Codigo = rs("Fuente")
      End If
      
    Case "ChequeBanco"
       If Not IsNull(rs("CodCuentas")) Then
       
'          FrmCheque.DBCodigo.Text = rs("CodCuentas")
          Codigo = rs("CodCuentas")
       End If
   
  
    Case "Departamento"
       If Not IsNull(rs("CodGrupo")) Then
      
          Codigo = rs("CodGrupo")

          

       End If
       
 Case "SolicitudPagos"
       PegarSolicitudPago
       
  Case "SolicitudCheques"
       PegarSolicitudCheque
    
 Case "Egreso"
    FrmEgresos.DtaCuentas.Refresh
    FrmEgresos.DBGTransacciones.Columns(0).Text = rs("CodCuentas")
     Criterio = "CodCuentas='" & FrmEgresos.DBGTransacciones.Columns(0).Text & "'"
      FrmEgresos.DtaCuentas.Recordset.Find (Criterio)
       If Not FrmEgresos.DtaCuentas.Recordset.EOF Then
         mes = Month(FrmEgresos.TxtFecha.Value)
         Año = Year(FrmEgresos.TxtFecha.Value)
         FechaIni = CDate("1/" & Month(FrmEgresos.TxtFecha.Value) & "/" & Year(FrmEgresos.TxtFecha.Value))
         FechaFin = DateSerial(Año, mes + 1, 1 - 1)
         NumFecha1 = CDate(FechaIni)
         NumFecha2 = CDate(FechaFin)
 
        FrmEgresos.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
        FrmEgresos.DtaConsulta.Refresh
         If Not FrmEgresos.DtaConsulta.Recordset.EOF Then
          FrmEgresos.TxtPeriodo.Text = FrmEgresos.DtaConsulta.Recordset("Periodo")
            NumeroPeriodo = FrmEgresos.DtaConsulta.Recordset("NPeriodo")
            If Val(FrmEgresos.TxtNTransacciones.Text) = 0 Then
                NumeroTransaccion = FrmEgresos.DtaConsulta.Recordset("NTransacciones")
            Else
                NumeroTransaccion = FrmEgresos.TxtNTransacciones.Text
            End If
            EstadoPeriodo = FrmEgresos.DtaConsulta.Recordset("EstadoPeriodo")
      
        '////////////Edito los datos del Periodo///////////
         If Val(FrmEgresos.TxtNTransacciones.Text) = 0 Then
          
          
'         FrmEgresos.'DtaConsulta.Recordset.Edit
         FrmEgresos.DtaConsulta.Recordset("NTransacciones") = FrmEgresos.DtaConsulta.Recordset("NTransacciones") + 1
         FrmEgresos.DtaConsulta.Recordset.Update
          NumeroTransaccion = FrmEgresos.DtaConsulta.Recordset("NTransacciones")
         FrmEgresos.TxtNTransacciones.Text = NumeroTransaccion
          '////////Edito los Datos de los indices de Transacciones//////
         
         FrmEgresos.DtaIndice.Recordset.AddNew
         FrmEgresos.DtaIndice.Recordset("FechaTransaccion") = Format(FrmEgresos.TxtFecha.Value, "dd/mm/yyyy")
         FrmEgresos.DtaIndice.Recordset("NumeroMovimiento") = NumeroTransaccion
         FrmEgresos.DtaIndice.Recordset("Fuente") = FrmEgresos.TxtFuente.Text
         FrmEgresos.DtaIndice.Recordset("NPeriodo") = NumeroPeriodo
         
      Criterio = "CodCuentas='" & FrmEgresos.DBGTransacciones.Columns(0).Text & "'"
       FrmEgresos.DtaCuentas.Recordset.Find (Criterio)
       If Not FrmEgresos.DtaCuentas.Recordset.EOF Then
        TipoMoneda = FrmEgresos.DtaCuentas.Recordset("TipoMoneda")

         Select Case TipoMoneda
            Case "Córdobas"
            
                      Fecha = FrmEgresos.TxtFecha.Value
                      Fechas = Format(FrmEgresos.TxtFecha.Value, "yyyy/mm/dd")
'                      FrmEgresos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmEgresos.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE (FechaTasas = '" & Format(FrmEgresos.TxtFecha.Value, "yyyymmdd") & "')"
                      FrmEgresos.DtaTasas.Refresh
                If Not FrmEgresos.DtaTasas.Recordset.EOF Then
                 MontoTasa = FrmEgresos.DtaTasas.Recordset("MontoCordobas")
                 If MontoTasa = 0 Then
                   MontoTasa = 1
                 End If
                 Select Case FrmEgresos.CmbMoneda.Text
                  Case "Córdobas"
                    FrmEgresos.DBGTransacciones.Columns(7).Text = 1
                  Case "Dólares"
                    FrmEgresos.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Libras"
                    FrmEgresos.DBGTransacciones.Columns(7).Text = MontoTasa
                   ' FrmEgresos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                    
                 End Select
                Else
                
                 FrmEgresos.DBGTransacciones.Columns(7).Text = 1
                End If
            
            Case "Dólares"
             Fecha = FrmEgresos.TxtFecha.Value
                      Fechas = Format(FrmEgresos.TxtFecha.Value, "yyyy/mm/dd")
'                      FrmEgresos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmEgresos.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas & "', 102))"
                      FrmEgresos.DtaTasas.Refresh
             If Not FrmEgresos.DtaTasas.Recordset.EOF Then
                MontoTasa = FrmEgresos.DtaTasas.Recordset("MontoCordobas")
                If MontoTasa = 0 Then
                  MontoTasa = 1
                End If
            
               Select Case FrmEgresos.CmbMoneda.Text
                  Case "Córdobas"
                    FrmEgresos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                  Case "Dólares"
                    FrmEgresos.DBGTransacciones.Columns(7).Text = 1
                  Case "Libras"
                    MontoTasa = FrmEgresos.DtaTasas.Recordset("MontoLibras")
                    FrmEgresos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)

                    
                 End Select
                Else
                  FrmEgresos.DBGTransacciones.Columns(7).Text = 1
               End If
            
            Case "Libras"
                      Fecha = FrmEgresos.TxtFecha.Value
                                            Fechas = Format(FrmEgresos.TxtFecha.Value, "yyyy/mm/dd")
'                      FrmEgresos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmEgresos.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas & "', 102))"
                      FrmEgresos.DtaTasas.Refresh
                If Not FrmEgresos.DtaTasas.Recordset.EOF Then
                MontoTasa = FrmEgresos.DtaTasas.Recordset("MontoLibras")
               Select Case FrmEgresos.CmbMoneda.Text
                  Case "Córdobas"
                    FrmEgresos.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Dólares"
                    FrmEgresos.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Libras"
                    FrmEgresos.DBGTransacciones.Columns(7).Text = 1

                    
                 End Select
                Else
                 FrmEgresos.DBGTransacciones.Columns(7).Text = 1
                End If
         
         End Select
       End If
         
         
         
         If FrmEgresos.CmbMoneda.Text = "Dólares" Then
            FrmEgresos.DtaIndice.Recordset("TipoMoneda") = "Dólares"
          Else
            FrmEgresos.DtaIndice.Recordset("TipoMoneda") = "Córdobas"
          End If
         
         FrmEgresos.DtaIndice.Recordset.Update
         Else
       Criterio = "CodCuentas='" & FrmEgresos.DBGTransacciones.Columns(0).Text & "'"
       FrmEgresos.DtaCuentas.Recordset.Find (Criterio)
       If Not FrmEgresos.DtaCuentas.Recordset.EOF Then
        TipoMoneda = FrmEgresos.DtaCuentas.Recordset("TipoMoneda")

         Select Case TipoMoneda
            Case "Córdobas"
                      Fecha = FrmEgresos.TxtFecha.Value
                       Fechas = Format(FrmEgresos.TxtFecha.Value, "yyyy/mm/dd")
'                      FrmEgresos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmEgresos.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas & "', 102))"
                      FrmEgresos.DtaTasas.Refresh
                If Not FrmEgresos.DtaTasas.Recordset.EOF Then
                 MontoTasa = FrmEgresos.DtaTasas.Recordset("MontoCordobas")
                 If MontoTasa = 0 Then
                   MontoTasa = 1
                 End If
                 Select Case FrmEgresos.CmbMoneda.Text
                  Case "Córdobas"
                    FrmEgresos.DBGTransacciones.Columns(7).Text = 1
                  Case "Dólares"
                    FrmEgresos.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Libras"
                    FrmEgresos.DBGTransacciones.Columns(7).Text = MontoTasa
                   ' FrmEgresos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                    
                 End Select
                Else
                
                 FrmEgresos.DBGTransacciones.Columns(7).Text = 1
                End If
            
            Case "Dólares"
             Fecha = FrmEgresos.TxtFecha.Value
             FrmEgresos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
             FrmEgresos.DtaTasas.Refresh
             If Not FrmEgresos.DtaTasas.Recordset.EOF Then
                MontoTasa = FrmEgresos.DtaTasas.Recordset("MontoCordobas")
                If MontoTasa = 0 Then
                  MontoTasa = 1
                End If
            
               Select Case FrmEgresos.CmbMoneda.Text
                  Case "Córdobas"
                    FrmEgresos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                  Case "Dólares"
                    FrmEgresos.DBGTransacciones.Columns(7).Text = 1
                  Case "Libras"
                    MontoTasa = FrmEgresos.DtaTasas.Recordset("MontoLibras")
                    FrmEgresos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)

                    
                 End Select
                Else
                  FrmEgresos.DBGTransacciones.Columns(7).Text = 1
               End If
            
            Case "Libras"
                      Fecha = FrmEgresos.TxtFecha.Value
                                            Fechas = Format(FrmEgresos.TxtFecha.Value, "yyyy/mm/dd")
'                      FrmEgresos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmEgresos.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas & "', 102))"
                      FrmEgresos.DtaTasas.Refresh
                If Not FrmEgresos.DtaTasas.Recordset.EOF Then
                MontoTasa = FrmEgresos.DtaTasas.Recordset("MontoLibras")
               Select Case FrmEgresos.CmbMoneda.Text
                  Case "Córdobas"
                    FrmEgresos.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Dólares"
                    FrmEgresos.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Libras"
                    FrmEgresos.DBGTransacciones.Columns(7).Text = 1

                    
                 End Select
                Else
                 FrmEgresos.DBGTransacciones.Columns(7).Text = 1
                End If
         
         End Select
       End If
           
         
         End If
        End If
 FrmSolicitudPagos.DtaCuentas.Refresh
  Criterio = "CodCuentas='" & FrmEgresos.DBCodigo.Text & "'"
  FrmSolicitudPagos.DtaCuentas.Recordset.Find (Criterio)
        
   TipoCuenta = FrmSolicitudPagos.DtaCuentas.Recordset("TipoCuenta")
   CodigoCuenta = FrmSolicitudPagos.DtaCuentas.Recordset("CodCuentas")
  If TipoCuenta = "Bancos" Then

   If Primero = True Then
     Primero = False
        Me.DtaConsulta.RecordSource = "SELECT NConsecutivoVoucher.CodCuenta, NConsecutivoVoucher.ConsecutivoVoucher, NConsecutivoVoucher.NPeriodo From NConsecutivoVoucher Where (((NConsecutivoVoucher.CodCuenta) = '" & CodigoCuenta & "') And ((NConsecutivoVoucher.NPeriodo) = " & NumeroPeriodo & "))"
        Me.DtaConsulta.Refresh
        If DtaConsulta.Recordset.EOF Then
           Me.DtaConsulta.Recordset.AddNew
             Me.DtaConsulta.Recordset("CodCuenta") = CodigoCuenta
             Me.DtaConsulta.Recordset("NPeriodo") = NumeroPeriodo
             Me.DtaConsulta.Recordset("ConsecutivoVoucher") = 1
           Me.DtaConsulta.Recordset.Update
           NumeroVoucher = 1
        Else
           'Me.'DtaConsulta.Recordset.Edit
            Me.DtaConsulta.Recordset("ConsecutivoVoucher") = Me.DtaConsulta.Recordset("ConsecutivoVoucher") + 1
           Me.DtaConsulta.Recordset.Update
         NumeroVoucher = Me.DtaConsulta.Recordset("ConsecutivoVoucher")
        End If
     Else
       ' FrmEgresos.DtaTransacciones.Recordset.MoveLast
     
     End If
        ConsecutivoVoucher = Month(FrmEgresos.TxtFecha.Value)
        If TipoCuenta = "Caja" Then
              numero = "CASH " & NumeroVoucher & "/" & ConsecutivoVoucher
        End If
        Select Case TipoMoneda
           Case "Córdobas"
            If TipoCuenta = "Bancos" Then
              numero = "BC " & NumeroVoucher & "/" & ConsecutivoVoucher
            End If
           Case "Dólares"
            If TipoCuenta = "Bancos" Then
              numero = "BD " & NumeroVoucher & "/" & ConsecutivoVoucher
            End If
        
         End Select
        
     End If
        ' Cadena = Mid(FrmEgresos.DBCodigo, 1, 1)
        ' Cadena = Cadena & "/" & NumeroTransaccion
        
   '///////////////////////////////////////////////////////////
   '//////CON ESTA CONSULTA BUSCO LA DESCRIPCION DE LA LINEA ANTERIOR//////////////////
   '/////////////////////////////////////////////////////////////////////////////////
   
            
            Sql = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta AS DescripcionCuentas, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, " & _
            "Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, " & _
            "Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, " & _
            "Transacciones.NumeroMovimiento , Periodos.Periodo " & _
            "FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo " & _
            "WHERE  (Transacciones.FechaTransaccion BETWEEN '" & Format(FrmEgresos.TxtFecha.Value, "yyyymmdd") & "' And '" & Format(FrmEgresos.TxtFecha.Value, "yyyymmdd") & "') AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ") " & _
            "ORDER BY Transacciones.NTransaccion"
              
            Me.DtaConsulta.RecordSource = Sql
            Me.DtaConsulta.Refresh
            If Not Me.DtaConsulta.Recordset.EOF Then
              Me.DtaConsulta.Recordset.MoveLast
              If Not IsNull(Me.DtaConsulta.Recordset("DescripcionMovimiento")) Then
                 DescripcionMovimiento = Me.DtaConsulta.Recordset("DescripcionMovimiento")
              End If
              If Not IsNull(Me.DtaConsulta.Recordset("Clave")) Then
                ClaveMovimiento = Me.DtaConsulta.Recordset("Clave")
              End If
            
            End If
        
        FrmEgresos.DBGTransacciones.Columns(3).Text = DescripcionMovimiento
        FrmEgresos.DBGTransacciones.Columns(2).Text = cadena
        If ClaveMovimiento = "" Then
         FrmEgresos.DBGTransacciones.Columns(6).Text = "Debito"
        Else
         FrmEgresos.DBGTransacciones.Columns(6).Text = ClaveMovimiento
        End If
        'FrmEgresos.DBGTransacciones.Columns(9).Locked = True
        FrmEgresos.DBGTransacciones.Columns(1).Text = rs("DescripcionCuentas")          'FrmEgresos.DtaCuentas.Recordset("DescripcionCuentas")
        FrmEgresos.DBGTransacciones.Columns(10).Text = Format(FrmEgresos.TxtFecha.Value, "dd/mm/yyyy")
        FrmEgresos.DBGTransacciones.Columns(11).Text = NumeroPeriodo
        FrmEgresos.DBGTransacciones.Columns(13).Text = FrmEgresos.TxtFuente.Text
        FrmEgresos.DBGTransacciones.Columns(14).Text = Format(FrmEgresos.TxtFecha.Value, "dd/mm/yyyy")
        FrmEgresos.DBGTransacciones.Columns(15).Text = NumeroTransaccion
         
'         For I = 2 To 5
'            If FrmEgresos.DBGTransacciones.Columns(I).Text = "" Then FrmEgresos.DBGTransacciones.Columns(I).Text = "-"
'        Next I
         
         'prueba de parche
'        FrmEgresos.DtaTransacciones.Refresh
'        inputFrmEgresos.DtaTransacciones.RecordSource
'         FrmEgresos.DBGTransacciones.Update
'
       Else
         MsgBox "La cuenta digitada no es correcta", vbCritical, "Sistema Contable"
         Exit Sub
       End If

    Case "Prorrateo2"
       If FrmProrrateo.TxtProrrateo.Text <> "" Then
         NumeroProrrateo = FrmProrrateo.TxtProrrateo.Text
         FrmProrrateo.AdoConsulta.RecordSource = "SELECT * From Prorrateo Where (NumeroProrrateo = " & NumeroProrrateo & " )"
         FrmProrrateo.AdoConsulta.Refresh
         If FrmProrrateo.AdoConsulta.Recordset.EOF Then
            FrmProrrateo.AdoProrrateo.Recordset.AddNew
             FrmProrrateo.AdoProrrateo.Recordset("NumeroProrrateo") = FrmProrrateo.TxtProrrateo.Text
            FrmProrrateo.AdoProrrateo.Recordset.Update
        
         End If
       Else

        MsgBox "Debe Crear primero el Prorrateo", vbCritical, "Sistema Contable"
        Exit Sub
       End If
       
        
       
         FrmProrrateo.TDBGridDestino.Columns(0).Text = FrmProrrateo.TxtProrrateo.Text
         FrmProrrateo.TDBGridDestino.Columns(1).Text = rs("CodCuentas")
         CodCuenta = rs("CodCuentas")
         FrmProrrateo.TDBGridDestino.Columns(2).Text = "DESTINO"
         FrmProrrateo.TDBGridDestino.Columns(3).Text = rs("DescripcionCuentas")
         FechaIni = Format(FrmProrrateo.DTPFechaIni.Value, "yyyy-mm-dd")
         FechaFin = Format(FrmProrrateo.DTPFechaFin.Value, "yyyy-mm-dd")
         FrmProrrateo.TDBGridDestino.Columns(4).Text = SaldoPeriodoCuentaDebito(FechaIni, FechaFin, CodCuenta)

         
    Case "Prorrateo"
       If FrmProrrateo.TxtProrrateo.Text <> "" Then
         NumeroProrrateo = FrmProrrateo.TxtProrrateo.Text
         FrmProrrateo.AdoConsulta.RecordSource = "SELECT * From Prorrateo Where (NumeroProrrateo = " & NumeroProrrateo & " )"
         FrmProrrateo.AdoConsulta.Refresh
         If FrmProrrateo.AdoConsulta.Recordset.EOF Then
            FrmProrrateo.AdoProrrateo.Recordset.AddNew
             FrmProrrateo.AdoProrrateo.Recordset("NumeroProrrateo") = FrmProrrateo.TxtProrrateo.Text
            FrmProrrateo.AdoProrrateo.Recordset.Update
        
         End If
       Else

        MsgBox "Debe Crear primero el Prorrateo", vbCritical, "Sistema Contable"
        Exit Sub
       End If
       
         FrmProrrateo.TDBGridOrigen.Columns(0).Text = FrmProrrateo.TxtProrrateo.Text
         FrmProrrateo.TDBGridOrigen.Columns(1).Text = rs("CodCuentas")
         CodCuenta = rs("CodCuentas")
         FrmProrrateo.TDBGridOrigen.Columns(2).Text = "ORIGEN"
         FrmProrrateo.TDBGridOrigen.Columns(3).Text = rs("DescripcionCuentas")
         FechaIni = Format(FrmProrrateo.DTPFechaIni.Value, "yyyy-mm-dd")
         FechaFin = Format(FrmProrrateo.DTPFechaFin.Value, "yyyy-mm-dd")
         FrmProrrateo.TDBGridOrigen.Columns(4).Text = SaldoPeriodoCuentaDebito(FechaIni, FechaFin, CodCuenta)
         
    
    Case "Auxiliar"
       If Not IsNull(rs("CodCuentas")) Then
          FrmAuxiliarCuentas.DBCliente.Text = rs("CodCuentas")
       End If
    Case "CuentaDepreciacion"
       If Not IsNull(rs("CodCuentas")) Then
          FrmActivoFijo.TxtDepreciacion = rs("CodCuentas")
       End If
       
    Case "CuentaGastos"
       If Not IsNull(rs("CodCuentas")) Then
          FrmActivoFijo.TxtGastos = rs("CodCuentas")
       End If
       
    Case "CuentaOriginal"
       If Not IsNull(rs("CodCuentas")) Then
          FrmActivoFijo.TxtCuentaOriginal = rs("CodCuentas")
       End If
       
    Case "Periodo"
       If Not IsNull(rs("CodCuentas")) Then
          FrmPeriodos.TxtContracuenta.Text = rs("CodCuentas")
          DescripcionContracuenta = rs("DescripcionCuentas")
          MonedaContracuenta = rs("TipoMoneda")
       End If
    Case "Cuenta"
       If Not IsNull(rs("CodCuentas")) Then
         If Not rs.EOF Then
          FrmCuentas.DBCliente.Text = rs("CodCuentas")
          FrmCuentas.TxtCodCuentas.Text = rs("CodCuentas")
          End If
       End If
       
    Case "CuentaFactura"
       If Not IsNull(rs("CodCuentas")) Then
          FrmTransacciones.TDBProveedor.Text = rs("CodCuentas")
          FrmTransacciones.LblNombres.Caption = rs("DescripcionCuentas")
       End If
    Case "CuentaActivoFijo"
       If Not IsNull(rs("CodCuenta")) Then
          FrmActivoFijo.DBCodigo.Text = rs("CodCuenta")
          
       End If
    Case "CuentaReportes"
       If Not IsNull(rs("CodCuentas")) Then
          FrmReportes.DBCodigo.Text = rs("CodCuentas")
       End If
       
    Case "CuentaReportes2"
       If Not IsNull(rs("CodCuentas")) Then
          FrmReportes.DBCodigoHasta.Text = rs("CodCuentas")
       End If
       
    Case "CuentaMayor"
       If Not IsNull(rs("CodCuentas")) Then
          CodigoCuenta = rs("CodCuentas")
       End If
    
    Case "Transferencia2"
       If Not IsNull(rs("CodCuentas")) Then
          FrmTransferencia.TxtDestino.Text = rs("CodCuentas")
       End If
    Case "Transferencia1"
       If Not IsNull(rs("CodCuentas")) Then
          FrmTransferencia.Txtorigen.Text = rs("CodCuentas")
       End If
    Case "ContratistasM"
       If Not IsNull(rs("CodCuentas")) Then
          FrmJustificacion.DBGMovimiento.Columns(3).Text = rs("CodCuentas")
       End If
       
     Case "MisCode"
       If Not IsNull(rs("CodCuentas")) Then
          FrmJustificacion.DBGMovimiento.Columns(5).Text = rs("CodCuentas")
       End If
  
    Case "Contratista"
   If Not IsNull(rs("Beneficiario")) Then
     FrmContactos.DBContratista.Text = rs("CodigoCuenta")
     Unload Me
   End If
  
  
  Case "ContratistaCheque"
   If Not IsNull(rs("Beneficiario")) Then
     FrmCheque.TxtNombre.Text = rs("Beneficiario")
     Unload Me
   End If
       
  Case "Cheque"
    FrmCheque.DtaCuentas.Refresh
    
    If rs.EOF Then
          Exit Sub
    End If
    FrmCheque.DBGTransacciones.Columns(0).Text = rs("CodCuentas")
     Criterio = "CodCuentas='" & FrmCheque.DBGTransacciones.Columns(0).Text & "'"
      FrmCheque.DtaCuentas.Recordset.Find (Criterio)
       If Not FrmCheque.DtaCuentas.Recordset.EOF Then
         mes = Month(FrmCheque.TxtFecha.Value)
         Año = Year(FrmCheque.TxtFecha.Value)
         FechaIni = CDate("1/" & Month(FrmCheque.TxtFecha.Value) & "/" & Year(FrmCheque.TxtFecha.Value))
         FechaFin = DateSerial(Año, mes + 1, 1 - 1)
         NumFecha1 = CDate(FechaIni)
         NumFecha2 = CDate(FechaFin)
 
        FrmCheque.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
        FrmCheque.DtaConsulta.Refresh
         If Not FrmCheque.DtaConsulta.Recordset.EOF Then
          FrmCheque.TxtPeriodo.Text = FrmCheque.DtaConsulta.Recordset("Periodo")
            NumeroPeriodo = FrmCheque.DtaConsulta.Recordset("NPeriodo")
            If Val(FrmCheque.TxtNTransacciones.Text) = 0 Then
                NumeroTransaccion = FrmCheque.DtaConsulta.Recordset("NTransacciones")
            Else
                NumeroTransaccion = FrmCheque.TxtNTransacciones.Text
            End If
            EstadoPeriodo = FrmCheque.DtaConsulta.Recordset("EstadoPeriodo")
      
        '////////////Edito los datos del Periodo///////////
         If Val(FrmCheque.TxtNTransacciones.Text) = 0 Then
          
          
'         FrmCheque.'DtaConsulta.Recordset.Edit
         FrmCheque.DtaConsulta.Recordset("NTransacciones") = FrmCheque.DtaConsulta.Recordset("NTransacciones") + 1
         FrmCheque.DtaConsulta.Recordset.Update
          NumeroTransaccion = FrmCheque.DtaConsulta.Recordset("NTransacciones")
         FrmCheque.TxtNTransacciones.Text = NumeroTransaccion
          '////////Edito los Datos de los indices de Transacciones//////
         
         FrmCheque.DtaIndice.Recordset.AddNew
         FrmCheque.DtaIndice.Recordset("FechaTransaccion") = Format(FrmCheque.TxtFecha.Value, "dd/mm/yyyy")
         FrmCheque.DtaIndice.Recordset("NumeroMovimiento") = NumeroTransaccion
         FrmCheque.DtaIndice.Recordset("Fuente") = FrmCheque.TxtFuente.Text
         FrmCheque.DtaIndice.Recordset("NPeriodo") = NumeroPeriodo
         
      Criterio = "CodCuentas='" & FrmCheque.DBGTransacciones.Columns(0).Text & "'"
       FrmCheque.DtaCuentas.Recordset.Find (Criterio)
       If Not FrmCheque.DtaCuentas.Recordset.EOF Then
        TipoMoneda = FrmCheque.DtaCuentas.Recordset("TipoMoneda")

         Select Case TipoMoneda
            Case "Córdobas"
            
                      Fecha = FrmCheque.TxtFecha.Value
                      Fechas = Format(FrmCheque.TxtFecha.Value, "yyyy/mm/dd")
'                      FrmCheque.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmCheque.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE (FechaTasas = '" & Format(FrmCheque.TxtFecha.Value, "yyyymmdd") & "')"
                      FrmCheque.DtaTasas.Refresh
                If Not FrmCheque.DtaTasas.Recordset.EOF Then
                 MontoTasa = FrmCheque.DtaTasas.Recordset("MontoCordobas")
                 If MontoTasa = 0 Then
                   MontoTasa = 1
                 End If
                 Select Case FrmCheque.CmbMoneda.Text
                  Case "Córdobas"
                    FrmCheque.DBGTransacciones.Columns(7).Text = 1
                  Case "Dólares"
                    FrmCheque.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Libras"
                    FrmCheque.DBGTransacciones.Columns(7).Text = MontoTasa
                   ' FrmCheque.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                    
                 End Select
                Else
                
                 FrmCheque.DBGTransacciones.Columns(7).Text = 1
                End If
            
            Case "Dólares"
             Fecha = FrmCheque.TxtFecha.Value
                      Fechas = Format(FrmCheque.TxtFecha.Value, "yyyy/mm/dd")
'                      FrmCheque.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmCheque.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas & "', 102))"
                      FrmCheque.DtaTasas.Refresh
             If Not FrmCheque.DtaTasas.Recordset.EOF Then
                MontoTasa = FrmCheque.DtaTasas.Recordset("MontoCordobas")
                If MontoTasa = 0 Then
                  MontoTasa = 1
                End If
            
               Select Case FrmCheque.CmbMoneda.Text
                  Case "Córdobas"
                    FrmCheque.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                  Case "Dólares"
                    FrmCheque.DBGTransacciones.Columns(7).Text = 1
                  Case "Libras"
                    MontoTasa = FrmCheque.DtaTasas.Recordset("MontoLibras")
                    FrmCheque.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)

                    
                 End Select
                Else
                  FrmCheque.DBGTransacciones.Columns(7).Text = 1
               End If
            
            Case "Libras"
                      Fecha = FrmCheque.TxtFecha.Value
                                            Fechas = Format(FrmCheque.TxtFecha.Value, "yyyy/mm/dd")
'                      FrmCheque.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmCheque.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas & "', 102))"
                      FrmCheque.DtaTasas.Refresh
                If Not FrmCheque.DtaTasas.Recordset.EOF Then
                MontoTasa = FrmCheque.DtaTasas.Recordset("MontoLibras")
               Select Case FrmCheque.CmbMoneda.Text
                  Case "Córdobas"
                    FrmCheque.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Dólares"
                    FrmCheque.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Libras"
                    FrmCheque.DBGTransacciones.Columns(7).Text = 1

                    
                 End Select
                Else
                 FrmCheque.DBGTransacciones.Columns(7).Text = 1
                End If
         
         End Select
       End If
         
         
         
         If FrmCheque.CmbMoneda.Text = "Dólares" Then
            FrmCheque.DtaIndice.Recordset("TipoMoneda") = "Dólares"
          Else
            FrmCheque.DtaIndice.Recordset("TipoMoneda") = "Córdobas"
          End If
         
         FrmCheque.DtaIndice.Recordset.Update
         Else
       Criterio = "CodCuentas='" & FrmCheque.DBGTransacciones.Columns(0).Text & "'"
       FrmCheque.DtaCuentas.Recordset.Find (Criterio)
       If Not FrmCheque.DtaCuentas.Recordset.EOF Then
        TipoMoneda = FrmCheque.DtaCuentas.Recordset("TipoMoneda")

         Select Case TipoMoneda
            Case "Córdobas"
                      Fecha = FrmCheque.TxtFecha.Value
                       Fechas = Format(FrmCheque.TxtFecha.Value, "yyyy/mm/dd")
'                      FrmCheque.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmCheque.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas & "', 102))"
                      FrmCheque.DtaTasas.Refresh
                If Not FrmCheque.DtaTasas.Recordset.EOF Then
                 MontoTasa = FrmCheque.DtaTasas.Recordset("MontoCordobas")
                 If MontoTasa = 0 Then
                   MontoTasa = 1
                 End If
                 Select Case FrmCheque.CmbMoneda.Text
                  Case "Córdobas"
                    FrmCheque.DBGTransacciones.Columns(7).Text = 1
                  Case "Dólares"
                    FrmCheque.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Libras"
                    FrmCheque.DBGTransacciones.Columns(7).Text = MontoTasa
                   ' FrmCheque.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                    
                 End Select
                Else
                
                 FrmCheque.DBGTransacciones.Columns(7).Text = 1
                End If
            
            Case "Dólares"
             Fecha = FrmCheque.TxtFecha.Value
             FrmCheque.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
             FrmCheque.DtaTasas.Refresh
             If Not FrmCheque.DtaTasas.Recordset.EOF Then
                MontoTasa = FrmCheque.DtaTasas.Recordset("MontoCordobas")
                If MontoTasa = 0 Then
                  MontoTasa = 1
                End If
            
               Select Case FrmCheque.CmbMoneda.Text
                  Case "Córdobas"
                    FrmCheque.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                  Case "Dólares"
                    FrmCheque.DBGTransacciones.Columns(7).Text = 1
                  Case "Libras"
                    MontoTasa = FrmCheque.DtaTasas.Recordset("MontoLibras")
                    FrmCheque.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)

                    
                 End Select
                Else
                  FrmCheque.DBGTransacciones.Columns(7).Text = 1
               End If
            
            Case "Libras"
                      Fecha = FrmCheque.TxtFecha.Value
                                            Fechas = Format(FrmCheque.TxtFecha.Value, "yyyy/mm/dd")
'                      FrmCheque.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmCheque.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas & "', 102))"
                      FrmCheque.DtaTasas.Refresh
                If Not FrmCheque.DtaTasas.Recordset.EOF Then
                MontoTasa = FrmCheque.DtaTasas.Recordset("MontoLibras")
               Select Case FrmCheque.CmbMoneda.Text
                  Case "Córdobas"
                    FrmCheque.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Dólares"
                    FrmCheque.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Libras"
                    FrmCheque.DBGTransacciones.Columns(7).Text = 1

                    
                 End Select
                Else
                 FrmCheque.DBGTransacciones.Columns(7).Text = 1
                End If
         
         End Select
       End If
           
         
         End If
        End If
 FrmCheque.DtaCuentas.Refresh
  Criterio = "CodCuentas='" & FrmCheque.DBCodigo.Text & "'"
  FrmCheque.DtaCuentas.Recordset.Find (Criterio)
        
   TipoCuenta = FrmCheque.DtaCuentas.Recordset("TipoCuenta")
   CodigoCuenta = FrmCheque.DtaCuentas.Recordset("CodCuentas")
  If TipoCuenta = "Bancos" Then

   If Primero = True Then
     Primero = False
        Me.DtaConsulta.RecordSource = "SELECT NConsecutivoVoucher.CodCuenta, NConsecutivoVoucher.ConsecutivoVoucher, NConsecutivoVoucher.NPeriodo From NConsecutivoVoucher Where (((NConsecutivoVoucher.CodCuenta) = '" & CodigoCuenta & "') And ((NConsecutivoVoucher.NPeriodo) = " & NumeroPeriodo & "))"
        Me.DtaConsulta.Refresh
        If DtaConsulta.Recordset.EOF Then
           Me.DtaConsulta.Recordset.AddNew
             Me.DtaConsulta.Recordset("CodCuenta") = CodigoCuenta
             Me.DtaConsulta.Recordset("NPeriodo") = NumeroPeriodo
             Me.DtaConsulta.Recordset("ConsecutivoVoucher") = 1
           Me.DtaConsulta.Recordset.Update
           NumeroVoucher = 1
        Else
           'Me.'DtaConsulta.Recordset.Edit
            Me.DtaConsulta.Recordset("ConsecutivoVoucher") = Me.DtaConsulta.Recordset("ConsecutivoVoucher") + 1
           Me.DtaConsulta.Recordset.Update
         NumeroVoucher = Me.DtaConsulta.Recordset("ConsecutivoVoucher")
        End If
     Else
       ' FrmCheque.DtaTransacciones.Recordset.MoveLast
     
     End If
        ConsecutivoVoucher = Month(FrmCheque.TxtFecha.Value)
        If TipoCuenta = "Caja" Then
              numero = "CASH " & NumeroVoucher & "/" & ConsecutivoVoucher
        End If
        Select Case TipoMoneda
           Case "Córdobas"
            If TipoCuenta = "Bancos" Then
              numero = "BC " & NumeroVoucher & "/" & ConsecutivoVoucher
            End If
           Case "Dólares"
            If TipoCuenta = "Bancos" Then
              numero = "BD " & NumeroVoucher & "/" & ConsecutivoVoucher
            End If
        
         End Select
        
     End If
        ' Cadena = Mid(FrmCheque.DBCodigo, 1, 1)
        ' Cadena = Cadena & "/" & NumeroTransaccion
        
   '///////////////////////////////////////////////////////////
   '//////CON ESTA CONSULTA BUSCO LA DESCRIPCION DE LA LINEA ANTERIOR//////////////////
   '/////////////////////////////////////////////////////////////////////////////////
   
            
            Sql = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta AS DescripcionCuentas, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, " & _
            "Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, " & _
            "Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, " & _
            "Transacciones.NumeroMovimiento , Periodos.Periodo " & _
            "FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo " & _
            "WHERE  (Transacciones.FechaTransaccion BETWEEN '" & Format(FrmCheque.TxtFecha.Value, "yyyymmdd") & "' And '" & Format(FrmCheque.TxtFecha.Value, "yyyymmdd") & "') AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ") " & _
            "ORDER BY Transacciones.NTransaccion"
              
            Me.DtaConsulta.RecordSource = Sql
            Me.DtaConsulta.Refresh
            If Not Me.DtaConsulta.Recordset.EOF Then
              Me.DtaConsulta.Recordset.MoveLast
              If Not IsNull(Me.DtaConsulta.Recordset("DescripcionMovimiento")) Then
                 DescripcionMovimiento = Me.DtaConsulta.Recordset("DescripcionMovimiento")
              End If
              If Not IsNull(Me.DtaConsulta.Recordset("Clave")) Then
                ClaveMovimiento = Me.DtaConsulta.Recordset("Clave")
              End If
            
            End If
        
        FrmCheque.DBGTransacciones.Columns(3).Text = DescripcionMovimiento
        FrmCheque.DBGTransacciones.Columns(2).Text = cadena
        If ClaveMovimiento = "" Then
         FrmCheque.DBGTransacciones.Columns(6).Text = "Debito"
        Else
         FrmCheque.DBGTransacciones.Columns(6).Text = ClaveMovimiento
        End If
        'FrmCheque.DBGTransacciones.Columns(9).Locked = True
        FrmCheque.DBGTransacciones.Columns(1).Text = rs("DescripcionCuentas")          'FrmCheque.DtaCuentas.Recordset("DescripcionCuentas")
        FrmCheque.DBGTransacciones.Columns(10).Text = Format(FrmCheque.TxtFecha.Value, "dd/mm/yyyy")
        FrmCheque.DBGTransacciones.Columns(11).Text = NumeroPeriodo
        FrmCheque.DBGTransacciones.Columns(13).Text = FrmCheque.TxtFuente.Text
        FrmCheque.DBGTransacciones.Columns(14).Text = Format(FrmCheque.TxtFecha.Value, "dd/mm/yyyy")
        FrmCheque.DBGTransacciones.Columns(15).Text = NumeroTransaccion
         
'         For I = 2 To 5
'            If FrmCheque.DBGTransacciones.Columns(I).Text = "" Then FrmCheque.DBGTransacciones.Columns(I).Text = "-"
'        Next I
         
         'prueba de parche
'        FrmCheque.DtaTransacciones.Refresh
'        inputfrmcheque.DtaTransacciones.RecordSource
'         FrmCheque.DBGTransacciones.Update
'
       Else
         MsgBox "La cuenta digitada no es correcta", vbCritical, "Sistema Contable"
         Exit Sub
       End If
       
       
   Case "AuxiliarTransacciones"
    TipoCuenta = rs("TipoCuenta")
    CodigoCuenta = rs("CodCuentas")
     FrmAuxiliarMovimientos.DBGTransacciones.Columns(0).Text = rs("CodCuentas")
     Criterio = "CodCuentas='" & FrmAuxiliarMovimientos.DBGTransacciones.Columns(0).Text & "'"
       FrmAuxiliarMovimientos.DtaCuentas.Recordset.Find (Criterio)
       If Not FrmAuxiliarMovimientos.DtaCuentas.Recordset.EOF Then
         FrmAuxiliarMovimientos.CmbMoneda.Enabled = False
         mes = Month(FrmAuxiliarMovimientos.TxtFecha.Value)
         Año = Year(FrmAuxiliarMovimientos.TxtFecha.Value)
         FechaIni = CDate("1/" & Month(FrmAuxiliarMovimientos.TxtFecha.Value) & "/" & Year(FrmAuxiliarMovimientos.TxtFecha.Value))
         FechaFin = DateSerial(Año, mes + 1, 1 - 1)
         NumFecha1 = FechaIni
         NumFecha2 = FechaFin
 
         FrmAuxiliarMovimientos.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
         FrmAuxiliarMovimientos.DtaConsulta.Refresh
         If Not FrmAuxiliarMovimientos.DtaConsulta.Recordset.EOF Then
           FrmAuxiliarMovimientos.TxtPeriodo.Text = FrmAuxiliarMovimientos.DtaConsulta.Recordset("Periodo")
            NumeroPeriodo = FrmAuxiliarMovimientos.DtaConsulta.Recordset("NPeriodo")
            If Val(FrmAuxiliarMovimientos.TxtNTransacciones.Text) = 0 Then
                NumeroTransaccion = FrmAuxiliarMovimientos.DtaConsulta.Recordset("NTransacciones")
            End If
            EstadoPeriodo = FrmAuxiliarMovimientos.DtaConsulta.Recordset("EstadoPeriodo")
      
        '////////////Edito los datos del Periodo///////////
         If Val(FrmAuxiliarMovimientos.TxtNTransacciones.Text) = 0 Then
          
          
          'FrmAuxiliarMovimientos.'DtaConsulta.Recordset.Edit
          FrmAuxiliarMovimientos.DtaConsulta.Recordset("NTransacciones") = FrmAuxiliarMovimientos.DtaConsulta.Recordset("NTransacciones") + 1
          FrmAuxiliarMovimientos.DtaConsulta.Recordset.Update
          NumeroTransaccion = FrmAuxiliarMovimientos.DtaConsulta.Recordset("NTransacciones")
          FrmAuxiliarMovimientos.TxtNTransacciones.Text = NumeroTransaccion
          '////////Edito los Datos de los indices de Transacciones//////
         
          FrmAuxiliarMovimientos.DtaIndice.Recordset.AddNew
          FrmAuxiliarMovimientos.DtaIndice.Recordset("FechaTransaccion") = Format(FrmAuxiliarMovimientos.TxtFecha.Value, "dd/mm/yyyy")
          FrmAuxiliarMovimientos.DtaIndice.Recordset("NumeroMovimiento") = NumeroTransaccion
          FrmAuxiliarMovimientos.DtaIndice.Recordset("Fuente") = FrmAuxiliarMovimientos.TxtFuente.Text
          FrmAuxiliarMovimientos.DtaIndice.Recordset("NPeriodo") = NumeroPeriodo
          FrmAuxiliarMovimientos.DtaIndice.Recordset.Update
         
         End If
        End If
       
        Criterio = "CodCuentas='" & FrmAuxiliarMovimientos.DBGTransacciones.Columns(0).Text & "'"
       FrmAuxiliarMovimientos.DtaCuentas.Recordset.Find (Criterio)
       If Not FrmAuxiliarMovimientos.DtaCuentas.Recordset.EOF Then
        TipoMoneda = FrmAuxiliarMovimientos.DtaCuentas.Recordset("TipoMoneda")


         Select Case TipoMoneda
            Case "Córdobas"
                      Fecha = FrmAuxiliarMovimientos.TxtFecha.Value
                      FrmAuxiliarMovimientos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmAuxiliarMovimientos.DtaTasas.Refresh
                If Not FrmAuxiliarMovimientos.DtaTasas.Recordset.EOF Then
                 MontoTasa = FrmAuxiliarMovimientos.DtaTasas.Recordset("MontoCordobas")
                 If MontoTasa = 0 Then
                   MontoTasa = 1
                 End If
                 Select Case FrmAuxiliarMovimientos.CmbMoneda.Text
                  Case "Córdobas"
                    FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).Text = 1
                  Case "Dólares"
                    FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Libras"
                    FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).Text = MontoTasa
                   ' FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                    
                 End Select
                Else
                
                 FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).Text = 1
                End If
            
            Case "Dólares"
             Fecha = FrmAuxiliarMovimientos.TxtFecha.Value
             FrmAuxiliarMovimientos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
             FrmAuxiliarMovimientos.DtaTasas.Refresh
             If Not FrmAuxiliarMovimientos.DtaTasas.Recordset.EOF Then
                MontoTasa = FrmAuxiliarMovimientos.DtaTasas.Recordset("MontoCordobas")
                If MontoTasa = 0 Then
                  MontoTasa = 1
                End If
            
               Select Case FrmAuxiliarMovimientos.CmbMoneda.Text
                  Case "Córdobas"
                    FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                  Case "Dólares"
                    FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).Text = 1
                  Case "Libras"
                    MontoTasa = FrmAuxiliarMovimientos.DtaTasas.Recordset("MontoLibras")
                    FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)

                    
                 End Select
                Else
                  FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).Text = 1
               End If
            
            Case "Libras"
                      Fecha = FrmAuxiliarMovimientos.TxtFecha.Value
                      FrmAuxiliarMovimientos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmAuxiliarMovimientos.DtaTasas.Refresh
                If Not FrmAuxiliarMovimientos.DtaTasas.Recordset.EOF Then
                MontoTasa = FrmAuxiliarMovimientos.DtaTasas.Recordset("MontoLibras")
               Select Case FrmAuxiliarMovimientos.CmbMoneda.Text
                  Case "Córdobas"
                    FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Dólares"
                    FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Libras"
                    FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).Text = 1

                    
                 End Select
                Else
                 FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).Text = 1
                End If
         
         End Select
       End If
       
        
   TipoCuenta = rs("TipoCuenta")
   CodigoCuenta = rs("CodCuentas")
  If TipoCuenta = "Bancos" Or TipoCuenta = "Caja" Then

   If Primero = True Then
     Primero = False
        Me.DtaConsulta.RecordSource = "SELECT NConsecutivoVoucher.CodCuenta, NConsecutivoVoucher.ConsecutivoVoucher, NConsecutivoVoucher.NPeriodo From NConsecutivoVoucher Where (((NConsecutivoVoucher.CodCuenta) = '" & CodigoCuenta & "') And ((NConsecutivoVoucher.NPeriodo) = " & NumeroPeriodo & "))"
        Me.DtaConsulta.Refresh
        If DtaConsulta.Recordset.EOF Then
           Me.DtaConsulta.Recordset.AddNew
             Me.DtaConsulta.Recordset("CodCuenta") = CodigoCuenta
             Me.DtaConsulta.Recordset("NPeriodo") = NumeroPeriodo
             Me.DtaConsulta.Recordset("ConsecutivoVoucher") = 1
           Me.DtaConsulta.Recordset.Update
           NumeroVoucher = 1
        Else
           'Me.'DtaConsulta.Recordset.Edit
            Me.DtaConsulta.Recordset("ConsecutivoVoucher") = Me.DtaConsulta.Recordset("ConsecutivoVoucher") + 1
           Me.DtaConsulta.Recordset.Update
         NumeroVoucher = Me.DtaConsulta.Recordset("ConsecutivoVoucher")
        End If
     Else
       ' FrmCheque.DtaTransacciones.Recordset.MoveLast
     
     End If
        ConsecutivoVoucher = Month(FrmTransacciones.TxtFecha.Value)
        If TipoCuenta = "Caja" Then
              numero = "CASH " & NumeroVoucher & "/" & ConsecutivoVoucher
        End If
        Select Case TipoMoneda
           Case "Córdobas"
            If TipoCuenta = "Bancos" Then
              numero = "BC " & NumeroVoucher & "/" & ConsecutivoVoucher
            End If
           Case "Dólares"
            If TipoCuenta = "Bancos" Then
              numero = "BD " & NumeroVoucher & "/" & ConsecutivoVoucher
            End If
        
         End Select
        
     End If

       
       

         FrmAuxiliarMovimientos.DBGTransacciones.Columns(2).Text = numero
         FrmAuxiliarMovimientos.DBGTransacciones.Columns(6).Text = "Debito"
         'FrmAuxiliarMovimientos.DBGTransacciones.Columns(9).Locked = True
         FrmAuxiliarMovimientos.DBGTransacciones.Columns(1).Text = FrmAuxiliarMovimientos.DtaCuentas.Recordset("DescripcionCuentas")
         FrmAuxiliarMovimientos.DBGTransacciones.Columns(10).Text = FrmAuxiliarMovimientos.TxtFecha.Value
         FrmAuxiliarMovimientos.DBGTransacciones.Columns(11).Text = NumeroPeriodo
         FrmAuxiliarMovimientos.DBGTransacciones.Columns(13).Text = FrmAuxiliarMovimientos.TxtFuente.Text
         FrmAuxiliarMovimientos.DBGTransacciones.Columns(14).Text = FrmAuxiliarMovimientos.TxtFecha.Value
         FrmAuxiliarMovimientos.DBGTransacciones.Columns(15).Text = NumeroTransaccion
         
         
       Else
         MsgBox "La cuenta digitada no es correcta", vbCritical, "Sistema Contable"
         Exit Sub
       End If
  
   Case "Transacciones"
    FrmTransacciones.DtaCuentas.Refresh
    
    If rs.EOF Then
      Exit Sub
    End If
    
    TipoCuenta = rs("TipoCuenta")
    CodigoCuenta = rs("CodCuentas")
     FrmTransacciones.DBGTransacciones.Columns(0).Text = rs("CodCuentas")
     Criterio = "CodCuentas='" & FrmTransacciones.DBGTransacciones.Columns(0).Text & "'"
       FrmTransacciones.DtaCuentas.Recordset.Find (Criterio)
       If Not FrmTransacciones.DtaCuentas.Recordset.EOF Then
         FrmTransacciones.CmbMoneda.Enabled = False
         FrmTransacciones.CmdBuscarEmpleado.Enabled = False
         mes = Month(FrmTransacciones.TxtFecha.Value)
         Año = Year(FrmTransacciones.TxtFecha.Value)
         FechaIni = CDate("1/" & Month(FrmTransacciones.TxtFecha.Value) & "/" & Year(FrmTransacciones.TxtFecha.Value))
         FechaFin = DateSerial(Año, mes + 1, 1 - 1)
'         NumFecha1 = FechaIni
'         NumFecha2 = FechaFin
 
         FrmTransacciones.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "'))"
'         InputBox "", "", FrmTransacciones.DtaConsulta.RecordSource
         FrmTransacciones.DtaConsulta.Refresh
         If Not FrmTransacciones.DtaConsulta.Recordset.EOF Then
           FrmTransacciones.TxtPeriodo.Text = FrmTransacciones.DtaConsulta.Recordset("Periodo")
            NumeroPeriodo = FrmTransacciones.DtaConsulta.Recordset("NPeriodo")
            If Val(FrmTransacciones.TxtNTransacciones.Text) = 0 Then
                NumeroTransaccion = FrmTransacciones.DtaConsulta.Recordset("NTransacciones")
            Else
                NumeroTransaccion = FrmTransacciones.TxtNTransacciones.Text
            End If
            EstadoPeriodo = FrmTransacciones.DtaConsulta.Recordset("EstadoPeriodo")
      
        '////////////Edito los datos del Periodo///////////
         If Val(FrmTransacciones.TxtNTransacciones.Text) = 0 Then
          
          
'          FrmTransacciones.'DtaConsulta.Recordset.Edit
          FrmTransacciones.DtaConsulta.Recordset("NTransacciones") = FrmTransacciones.DtaConsulta.Recordset("NTransacciones") + 1
          FrmTransacciones.DtaConsulta.Recordset.Update
          NumeroTransaccion = FrmTransacciones.DtaConsulta.Recordset("NTransacciones")
          FrmTransacciones.TxtNTransacciones.Text = NumeroTransaccion
          
          '//////////////////////////////////////////////////////////////////////////////////////////////////////
          '////////////////////////Edito los Datos de los indices de Transacciones//////////////////////////////
          '//////////////////////////////////////////////////////////////////////////////////////////////////////
         
          FrmTransacciones.DtaIndice.Recordset.AddNew
          FrmTransacciones.DtaIndice.Recordset("FechaTransaccion") = Format(FrmTransacciones.TxtFecha.Value, "dd/mm/yyyy")
          FrmTransacciones.DtaIndice.Recordset("NumeroMovimiento") = NumeroTransaccion
          FrmTransacciones.DtaIndice.Recordset("Fuente") = FrmTransacciones.TxtFuente.Text
          FrmTransacciones.DtaIndice.Recordset("NPeriodo") = NumeroPeriodo
          FrmTransacciones.DtaIndice.Recordset("TipoMoneda") = FrmTransacciones.CmbMoneda.Text
          
     
          
          
          
          
          
          
          FrmTransacciones.DtaIndice.Recordset.Update
         
         End If
        End If
       
        Criterio = "CodCuentas='" & FrmTransacciones.DBGTransacciones.Columns(0).Text & "'"
       FrmTransacciones.DtaCuentas.Recordset.Find (Criterio)
       If Not FrmTransacciones.DtaCuentas.Recordset.EOF Then
        TipoMoneda = FrmTransacciones.DtaCuentas.Recordset("TipoMoneda")


         Select Case TipoMoneda
            Case "Córdobas"
                      Fecha = FrmTransacciones.TxtFecha.Value
                      'FrmTransacciones.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmTransacciones.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) ='" & Format(FrmTransacciones.TxtFecha.Value, "yyyymmdd") & "'))"
                      FrmTransacciones.DtaTasas.Refresh
                If Not FrmTransacciones.DtaTasas.Recordset.EOF Then
                 MontoTasa = FrmTransacciones.DtaTasas.Recordset("MontoCordobas")
                 If MontoTasa = 0 Then
                   MontoTasa = 1
                 End If
                 Select Case FrmTransacciones.CmbMoneda.Text
                  Case "Córdobas"
                    FrmTransacciones.DBGTransacciones.Columns("TCambio").Text = 1
                  Case "Dólares"
                    FrmTransacciones.DBGTransacciones.Columns("TCambio").Text = MontoTasa
                  Case "Libras"
                    FrmTransacciones.DBGTransacciones.Columns("TCambio").Text = MontoTasa
                   ' FrmTransacciones.DBGTransacciones.Columns("TCambio").Text = (1 / MontoTasa)
                    
                 End Select
                Else
                
                 FrmTransacciones.DBGTransacciones.Columns("TCambio").Text = 1
                End If
            
            Case "Dólares"
             Fecha = FrmTransacciones.TxtFecha.Value
             FrmTransacciones.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) ='" & Format(FrmTransacciones.TxtFecha.Value, "yyyymmdd") & "'))" 'jp feb
             FrmTransacciones.DtaTasas.Refresh
             If Not FrmTransacciones.DtaTasas.Recordset.EOF Then
                MontoTasa = FrmTransacciones.DtaTasas.Recordset("MontoCordobas")
                If MontoTasa = 0 Then
                  MontoTasa = 1
                End If
            
               Select Case FrmTransacciones.CmbMoneda.Text
                  Case "Córdobas"
                    FrmTransacciones.DBGTransacciones.Columns("TCambio").Text = (1 / MontoTasa)
                  Case "Dólares"
                    FrmTransacciones.DBGTransacciones.Columns("TCambio").Text = 1
                  Case "Libras"
                    MontoTasa = FrmTransacciones.DtaTasas.Recordset("MontoLibras")
                    FrmTransacciones.DBGTransacciones.Columns("TCambio").Text = (1 / MontoTasa)

                    
                 End Select
                Else
                  FrmTransacciones.DBGTransacciones.Columns("TCambio").Text = 1
               End If
            
            Case "Libras"
                      Fecha = FrmTransacciones.TxtFecha.Value
                      FrmTransacciones.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmTransacciones.DtaTasas.Refresh
                If Not FrmTransacciones.DtaTasas.Recordset.EOF Then
                MontoTasa = FrmTransacciones.DtaTasas.Recordset("MontoLibras")
               Select Case FrmTransacciones.CmbMoneda.Text
                  Case "Córdobas"
                    FrmTransacciones.DBGTransacciones.Columns("TCambio").Text = MontoTasa
                  Case "Dólares"
                    FrmTransacciones.DBGTransacciones.Columns("TCambio").Text = MontoTasa
                  Case "Libras"
                    FrmTransacciones.DBGTransacciones.Columns("TCambio").Text = 1

                    
                 End Select
                Else
                 FrmTransacciones.DBGTransacciones.Columns("TCambio").Text = 1
                End If
         
         End Select
       End If
       
        
   TipoCuenta = rs("TipoCuenta")
   CodigoCuenta = rs("CodCuentas")
  If TipoCuenta = "Bancos" Or TipoCuenta = "Caja" Then

   If Primero = True Then
     Primero = False
        Me.DtaConsulta.RecordSource = "SELECT NConsecutivoVoucher.CodCuenta, NConsecutivoVoucher.ConsecutivoVoucher, NConsecutivoVoucher.NPeriodo From NConsecutivoVoucher Where (((NConsecutivoVoucher.CodCuenta) = '" & CodigoCuenta & "') And ((NConsecutivoVoucher.NPeriodo) = " & NumeroPeriodo & "))"
        Me.DtaConsulta.Refresh
        If DtaConsulta.Recordset.EOF Then
           Me.DtaConsulta.Recordset.AddNew
             Me.DtaConsulta.Recordset("CodCuenta") = CodigoCuenta
             Me.DtaConsulta.Recordset("NPeriodo") = NumeroPeriodo
             Me.DtaConsulta.Recordset("ConsecutivoVoucher") = 1
           Me.DtaConsulta.Recordset.Update
           NumeroVoucher = 1
        Else
           'Me.'DtaConsulta.Recordset.Edit
            Me.DtaConsulta.Recordset("ConsecutivoVoucher") = Me.DtaConsulta.Recordset("ConsecutivoVoucher") + 1
           Me.DtaConsulta.Recordset.Update
         NumeroVoucher = Me.DtaConsulta.Recordset("ConsecutivoVoucher")
        End If
     Else
       ' FrmCheque.DtaTransacciones.Recordset.MoveLast
     
     End If
     
         ConsecutivoVoucher = Month(FrmTransacciones.TxtFecha.Value)
        If TipoCuenta = "Caja" Then
              numero = "CASH " & NumeroVoucher & "/" & ConsecutivoVoucher
        End If
        Select Case TipoMoneda
           Case "Córdobas"
            If TipoCuenta = "Bancos" Then
              numero = "BC " & NumeroVoucher & "/" & ConsecutivoVoucher
            End If
           Case "Dólares"
            If TipoCuenta = "Bancos" Then
              numero = "BD " & NumeroVoucher & "/" & ConsecutivoVoucher
            End If
        
         End Select
        
     End If

       
    '///////////////////////////////////////////////////////////
   '//////CON ESTA CONSULTA BUSCO LA DESCRIPCION DE LA LINEA ANTERIOR//////////////////
   '/////////////////////////////////////////////////////////////////////////////////
   
            
            Sql = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta AS DescripcionCuentas, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, " & _
            "Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, " & _
            "Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, " & _
            "Transacciones.NumeroMovimiento , Periodos.Periodo " & _
            "FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo " & _
            "WHERE  (Transacciones.FechaTransaccion BETWEEN '" & Format(FrmTransacciones.TxtFecha.Value, "yyyymmdd") & "' And '" & Format(FrmTransacciones.TxtFecha.Value, "yyyymmdd") & "') AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ") " & _
            "ORDER BY Transacciones.NTransaccion"
              
            FrmTransacciones.DtaConsulta.RecordSource = Sql
            FrmTransacciones.DtaConsulta.Refresh
            If Not FrmTransacciones.DtaConsulta.Recordset.EOF Then
              FrmTransacciones.DtaConsulta.Recordset.MoveLast
              If Not IsNull(FrmTransacciones.DtaConsulta.Recordset("DescripcionMovimiento")) Then
                 DescripcionMovimiento = FrmTransacciones.DtaConsulta.Recordset("DescripcionMovimiento")
              End If
              If Not IsNull(FrmTransacciones.DtaConsulta.Recordset("Clave")) Then
                ClaveMovimiento = FrmTransacciones.DtaConsulta.Recordset("Clave")
              End If
            End If
       
         FrmTransacciones.DBGTransacciones.Columns(3).Text = DescripcionMovimiento
         FrmTransacciones.DBGTransacciones.Columns(2).Text = numero
         If ClaveMovimiento = "" Then
          FrmTransacciones.DBGTransacciones.Columns(6).Text = "Debito"
         Else
          FrmTransacciones.DBGTransacciones.Columns(6).Text = ClaveMovimiento
         End If
         
         

         FrmTransacciones.DBGTransacciones.Columns(1).Text = FrmTransacciones.DtaCuentas.Recordset("DescripcionCuentas")
         FrmTransacciones.DBGTransacciones.Columns(10).Text = Format(FrmTransacciones.TxtFecha.Value, "dd/mm/yyyy")
         FrmTransacciones.DBGTransacciones.Columns(11).Text = NumeroPeriodo
         FrmTransacciones.DBGTransacciones.Columns(13).Text = FrmTransacciones.TxtFuente.Text
         FrmTransacciones.DBGTransacciones.Columns(14).Text = Format(FrmTransacciones.TxtFecha.Value, "dd/mm/yyyy")
         FrmTransacciones.DBGTransacciones.Columns(15).Text = NumeroTransaccion
         
         
'     FrmTransacciones.DBGTransacciones.Columns("CodCuentas").Button = True
'     FrmTransacciones.DBGTransacciones.Columns(0).Width = 1500
'     FrmTransacciones.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
'     FrmTransacciones.DBGTransacciones.Columns(6).Button = True
'     FrmTransacciones.DBGTransacciones.Columns(6).Locked = True
'
'     FrmTransacciones.DBGTransacciones.Columns(2).Width = 1000
'     FrmTransacciones.DBGTransacciones.Columns(3).Caption = "Descripcion"
'     FrmTransacciones.DBGTransacciones.Columns(4).Width = 1000
'     FrmTransacciones.DBGTransacciones.Columns(4).Button = True
'     FrmTransacciones.DBGTransacciones.Columns(5).Width = 1000
'     FrmTransacciones.DBGTransacciones.Columns(6).Width = 800
'     FrmTransacciones.DBGTransacciones.Columns(7).Width = 1200
'     FrmTransacciones.DBGTransacciones.Columns(7).NumberFormat = "##,##0.000000"
'     FrmTransacciones.DBGTransacciones.Columns(8).Width = 1200
'     FrmTransacciones.DBGTransacciones.Columns(8).NumberFormat = "##,##0.00"
'     FrmTransacciones.DBGTransacciones.Columns(9).Width = 1200
'     FrmTransacciones.DBGTransacciones.Columns(9).NumberFormat = "##,##0.00"
'     FrmTransacciones.DBGTransacciones.Columns(10).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(11).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(12).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(13).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(14).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(15).Visible = False
''     FrmTransacciones.DBGTransacciones.Columns(16).Visible = False
''     FrmTransacciones.DBGTransacciones.Columns(17).Visible = False
''     FrmTransacciones.DBGTransacciones.Columns(18).Visible = False
''     FrmTransacciones.DBGTransacciones.Columns(19).Visible = False
''     FrmTransacciones.DBGTransacciones.Columns(20).Visible = False
''     FrmTransacciones.DBGTransacciones.Columns(21).Visible = False
''     FrmTransacciones.DBGTransacciones.Columns(22).Visible = False
 
 
  FrmTransacciones.DBGTransacciones.Columns("CodCuentas").Width = 1500  '0
  FrmTransacciones.DBGTransacciones.Columns("CodCuentas").Button = True  '0
  FrmTransacciones.DBGTransacciones.Columns("NombreCuenta").Locked = True '1
  FrmTransacciones.DBGTransacciones.Columns("NombreCuenta").Locked = True '1
  FrmTransacciones.DBGTransacciones.Columns("VoucherNo").Width = 1000 '2
  FrmTransacciones.DBGTransacciones.Columns("VoucherNo").Caption = "Voucher/Dpto" '2
  FrmTransacciones.DBGTransacciones.Columns("VoucherNo").Width = 1100 '2
  FrmTransacciones.DBGTransacciones.Columns("DescripcionMovimiento").Caption = "Descripcion" '3
  FrmTransacciones.DBGTransacciones.Columns("FacturaNo").Width = 1000 '4
  FrmTransacciones.DBGTransacciones.Columns("FacturaNo").Width = 1000 '4
  FrmTransacciones.DBGTransacciones.Columns("FacturaNo").Button = True '4
  FrmTransacciones.DBGTransacciones.Columns("ChequeNo").Width = 1000 '5
  FrmTransacciones.DBGTransacciones.Columns("ChequeNo").Caption = "Cheq/Rec" '5
  FrmTransacciones.DBGTransacciones.Columns("Clave").Button = True '6
  FrmTransacciones.DBGTransacciones.Columns("Clave").Locked = True '6
  FrmTransacciones.DBGTransacciones.Columns("Clave").Width = 800   '6
  FrmTransacciones.DBGTransacciones.Columns("TCambio").Caption = "Tasa Cambio"         '7
  FrmTransacciones.DBGTransacciones.Columns("TCambio").Locked = True                   '7
  FrmTransacciones.DBGTransacciones.Columns("TCambio").NumberFormat = "##,##0.000000"  '7
  FrmTransacciones.DBGTransacciones.Columns("TCambio").Width = 1200 '8
  FrmTransacciones.DBGTransacciones.Columns("TCambio").Locked = True '8
  FrmTransacciones.DBGTransacciones.Columns("Debito").Width = 1200   '9
  FrmTransacciones.DBGTransacciones.Columns("Debito").NumberFormat = "##,##0.00" '9
  FrmTransacciones.DBGTransacciones.Columns("Credito").Width = 1200 '10
  FrmTransacciones.DBGTransacciones.Columns("Credito").NumberFormat = "##,##0.00" '10
  FrmTransacciones.DBGTransacciones.Columns("FechaTransaccion").Visible = False  '11
  FrmTransacciones.DBGTransacciones.Columns("NPeriodo").Visible = False '12
  FrmTransacciones.DBGTransacciones.Columns("NTransaccion").Visible = False  '13
  FrmTransacciones.DBGTransacciones.Columns("Fuente").Visible = False  '14
  FrmTransacciones.DBGTransacciones.Columns("FechaTasas").Visible = False  '15
  FrmTransacciones.DBGTransacciones.Columns("NumeroMovimiento").Visible = False '16
  FrmTransacciones.DBGTransacciones.Columns("Periodo").Visible = False '17
  FrmTransacciones.DBGTransacciones.Columns("FechaDescuento").Visible = False  '18
  FrmTransacciones.DBGTransacciones.Columns("DescuentoDisponible").Visible = False  '19
  FrmTransacciones.DBGTransacciones.Columns("FechaVence").Visible = False '20
  FrmTransacciones.DBGTransacciones.Columns("CodCuentaProveedor").Visible = False '21
  FrmTransacciones.DBGTransacciones.Columns("TipoFactura").Visible = False '22
  FrmTransacciones.DBGTransacciones.Columns("NTransaccion").Visible = False '23
       
       
       
       
       Else
         MsgBox "La cuenta digitada no es correcta", vbCritical, "Sistema Contable"
         Exit Sub
       End If
       
   Case "NTransacciones"
  If Not Me.DbgrProducto.Columns(2).Text = "" Then
    Descripcion = Me.DbgrProducto.Columns(2).Text
   If Descripcion = "*****CANCELADO*****" Then
     MsgBox " Este movimiento esta Cancelado", vbCritical, "Sistema Contable"
     Exit Sub
   End If
  End If
   
   
 FrmTransacciones.CmbMoneda.Enabled = False
   
 mes = Month(FrmTransacciones.TxtFecha.Value)
 Año = Year(FrmTransacciones.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(FrmTransacciones.TxtFecha.Value) & "/" & Year(FrmTransacciones.TxtFecha.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
' NumFecha1 = FechaIni
' NumFecha2 = FechaFin
 
 FrmTransacciones.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,  Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones WHERE Transacciones.FechaTransaccion>='" & Format(FechaIni, "yyyymmdd") & "' And transacciones.fechatransaccion<='" & Format(FechaFin, "yyyymmdd") & "' ORDER BY Transacciones.NumeroMovimiento"
 FrmTransacciones.DtaConsulta.Refresh
 
 If Not FrmTransacciones.DtaConsulta.Recordset.EOF Then
   NumeroTransaccion = rs("NumeroMovimiento")
   FrmTransacciones.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,  Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo,FechaDescuento,DescuentoDisponible,FechaVence,Transacciones.CodCuentaProveedor,Transacciones.TipoFactura FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE Transacciones.FechaTransaccion>='" & Format(NumFecha1, "yyyymmdd") & "' And Transacciones.FechaTransaccion<='" & Format(NumFecha2, "yyyymmdd") & "' AND Transacciones.NumeroMovimiento=" & NumeroTransaccion & " ORDER BY Transacciones.NTransaccion"
   FrmTransacciones.DtaTransacciones.Refresh
'   InputBox "", "", FrmTransacciones.DtaTransacciones.RecordSource
   If Not FrmTransacciones.DtaTransacciones.Recordset.EOF Then
     
     FrmTransacciones.TxtFecha.Value = FrmTransacciones.DtaTransacciones.Recordset("FechaTransaccion")
     FrmTransacciones.TxtPeriodo.Text = FrmTransacciones.DtaTransacciones.Recordset("Periodo")
     FrmTransacciones.TxtNTransacciones.Text = FrmTransacciones.DtaTransacciones.Recordset("NumeroMovimiento")
     NumeroTransaccion = FrmTransacciones.DtaTransacciones.Recordset("NumeroMovimiento")
     FrmTransacciones.TxtFuente.Text = FrmTransacciones.DtaTransacciones.Recordset("Fuente")

      FrmTransacciones.DtaConsulta.RecordSource = "SELECT IndiceTransaccion.TipoMoneda,IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente, IndiceTransaccion.Fuente From IndiceTransaccion WHERE IndiceTransaccion.FechaTransaccion>='" & Format(NumFecha1, "yyyymmdd") & "' And indicetransaccion.fechatransaccion<='" & Format(NumFecha2, "yyyymmdd") & "' AND IndiceTransaccion.NumeroMovimiento= " & NumeroTransaccion
      FrmTransacciones.DtaConsulta.Refresh
      If Not FrmTransacciones.DtaConsulta.Recordset.EOF Then
        If Not IsNull(FrmTransacciones.DtaConsulta.Recordset("TipoMoneda")) Then
            FrmTransacciones.CmbMoneda.Text = FrmTransacciones.DtaConsulta.Recordset("TipoMoneda")
        Else
            FrmTransacciones.CmbMoneda.Text = ""
        End If
      End If
     
     
     '//////Sumo los Totales/////////////////////
    Debito = 0
    Credito = 0
    TotalDebito = 0
    TotalCredito = 0
      NumFecha1 = FrmTransacciones.TxtFecha.Value
      Fechas = Format(FrmTransacciones.TxtFecha.Value, "yyyy/mm/dd")
      NMovimiento = Val(FrmTransacciones.TxtNTransacciones)
       FrmTransacciones.DtaConsulta.RecordSource = "SELECT FechaTransaccion, CodCuentas, NumeroMovimiento, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, TCambio, Debito, Credito From Transacciones WHERE     (FechaTransaccion = CONVERT(DATETIME, '" & Fechas & "', 102)) AND (NumeroMovimiento = " & NMovimiento & ")"
       
      FrmTransacciones.DtaConsulta.Refresh
      Do While Not FrmTransacciones.DtaConsulta.Recordset.EOF
       If Not IsNull(FrmTransacciones.DtaConsulta.Recordset("Debito")) Then
       Debito = FrmTransacciones.DtaConsulta.Recordset("Debito")
       End If
       If Not IsNull(Credito = FrmTransacciones.DtaConsulta.Recordset("Credito")) Then
        Credito = FrmTransacciones.DtaConsulta.Recordset("Credito")
       End If
       TotalDebito = TotalDebito + Debito
       TotalCredito = TotalCredito + Credito
       FrmTransacciones.DtaConsulta.Recordset.MoveNext
      Loop
    FrmTransacciones.TxtCredito.Text = Format(TotalCredito, "##,##0.00")
    FrmTransacciones.TxtDebito.Text = Format(TotalDebito, "##,##0.00")
    FrmTransacciones.TxtDiferencia.Text = Format(TotalDebito - TotalCredito, "##,##0.00")
        
  
  FrmTransacciones.DBGTransacciones.Columns("CodCuentas").Width = 1500  '0
  FrmTransacciones.DBGTransacciones.Columns("CodCuentas").Button = True  '0
  FrmTransacciones.DBGTransacciones.Columns("NombreCuenta").Locked = True '1
  FrmTransacciones.DBGTransacciones.Columns("NombreCuenta").Locked = True '1
  FrmTransacciones.DBGTransacciones.Columns("VoucherNo").Width = 1000 '2
  FrmTransacciones.DBGTransacciones.Columns("VoucherNo").Caption = "Voucher/Dpto" '2
  FrmTransacciones.DBGTransacciones.Columns("VoucherNo").Width = 1100 '2
  FrmTransacciones.DBGTransacciones.Columns("DescripcionMovimiento").Caption = "Descripcion" '3
  FrmTransacciones.DBGTransacciones.Columns("FacturaNo").Width = 1000 '4
  FrmTransacciones.DBGTransacciones.Columns("FacturaNo").Width = 1000 '4
  FrmTransacciones.DBGTransacciones.Columns("FacturaNo").Button = True '4
  FrmTransacciones.DBGTransacciones.Columns("ChequeNo").Width = 1000
  FrmTransacciones.DBGTransacciones.Columns("ChequeNo").Caption = "Cheq/Rec"
  FrmTransacciones.DBGTransacciones.Columns("Clave").Button = True
  FrmTransacciones.DBGTransacciones.Columns("Clave").Locked = True
  FrmTransacciones.DBGTransacciones.Columns("Clave").Width = 800
  FrmTransacciones.DBGTransacciones.Columns("TCambio").Caption = "Tasa Cambio"
  FrmTransacciones.DBGTransacciones.Columns("TCambio").Locked = True
  FrmTransacciones.DBGTransacciones.Columns("TCambio").NumberFormat = "##,##0.000000"
  FrmTransacciones.DBGTransacciones.Columns("TCambio").Width = 1200
  FrmTransacciones.DBGTransacciones.Columns("TCambio").Locked = True
  FrmTransacciones.DBGTransacciones.Columns("Debito").Width = 1200
  FrmTransacciones.DBGTransacciones.Columns("Debito").NumberFormat = "##,##0.00"
  FrmTransacciones.DBGTransacciones.Columns("Credito").Width = 1200
  FrmTransacciones.DBGTransacciones.Columns("Credito").NumberFormat = "##,##0.00"
  FrmTransacciones.DBGTransacciones.Columns("FechaTransaccion").Visible = False
  FrmTransacciones.DBGTransacciones.Columns("NPeriodo").Visible = False
  FrmTransacciones.DBGTransacciones.Columns("NTransaccion").Visible = False
  FrmTransacciones.DBGTransacciones.Columns("Fuente").Visible = False
  FrmTransacciones.DBGTransacciones.Columns("FechaTasas").Visible = False
  FrmTransacciones.DBGTransacciones.Columns("NumeroMovimiento").Visible = False
  FrmTransacciones.DBGTransacciones.Columns("Periodo").Visible = False
  FrmTransacciones.DBGTransacciones.Columns("FechaDescuento").Visible = False
  FrmTransacciones.DBGTransacciones.Columns("DescuentoDisponible").Visible = False
  FrmTransacciones.DBGTransacciones.Columns("FechaVence").Visible = False
  FrmTransacciones.DBGTransacciones.Columns("CodCuentaProveedor").Visible = False
  FrmTransacciones.DBGTransacciones.Columns("TipoFactura").Visible = False
  FrmTransacciones.DBGTransacciones.Columns("NTransaccion").Visible = False
       
       
     FrmTransacciones.TxtFecha.Enabled = False
     FrmTransacciones.TxtPeriodo.Enabled = False
     FrmTransacciones.TxtFuente.Enabled = False
     FrmTransacciones.TxtNTransacciones.Enabled = False
     
     
     End If
       
'      FrmTransacciones.DBGTransacciones.Columns(0).Button = True
'     FrmTransacciones.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
'     FrmTransacciones.DBGTransacciones.Columns(6).Button = True
'     FrmTransacciones.DBGTransacciones.Columns(6).Locked = True
'     FrmTransacciones.DBGTransacciones.Columns(0).Width = 1500
'     FrmTransacciones.DBGTransacciones.Columns(2).Width = 1100
'     FrmTransacciones.DBGTransacciones.Columns(2).Caption = "Voucher/Dpto"
'     FrmTransacciones.DBGTransacciones.Columns(3).Caption = "Descripcion"
'     FrmTransacciones.DBGTransacciones.Columns(4).Width = 1000
'     FrmTransacciones.DBGTransacciones.Columns(4).Button = True
'     FrmTransacciones.DBGTransacciones.Columns(5).Width = 1000
'     FrmTransacciones.DBGTransacciones.Columns(6).Width = 800
'     FrmTransacciones.DBGTransacciones.Columns(7).Width = 1200
'     FrmTransacciones.DBGTransacciones.Columns("TCambio").NumberFormat = "##,##0.000000"
'     FrmTransacciones.DBGTransacciones.Columns(8).Width = 1200
'     FrmTransacciones.DBGTransacciones.Columns(8).NumberFormat = "##,##0.00"
'     FrmTransacciones.DBGTransacciones.Columns(9).Width = 1200
'     FrmTransacciones.DBGTransacciones.Columns(9).NumberFormat = "##,##0.00"
'     FrmTransacciones.DBGTransacciones.Columns(10).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(11).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(12).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(13).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(14).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(15).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(16).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(17).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(18).Visible = False
'     FrmTransacciones.DBGTransacciones.Columns(19).Visible = False

        FrmTransacciones.DBGTransacciones.Columns("CodCuentas").Width = 1500  '0
        FrmTransacciones.DBGTransacciones.Columns("CodCuentas").Button = True  '0
        FrmTransacciones.DBGTransacciones.Columns("NombreCuenta").Locked = True '1
        FrmTransacciones.DBGTransacciones.Columns("NombreCuenta").Locked = True '1
        FrmTransacciones.DBGTransacciones.Columns("VoucherNo").Width = 1000 '2
        FrmTransacciones.DBGTransacciones.Columns("VoucherNo").Caption = "Voucher/Dpto" '2
        FrmTransacciones.DBGTransacciones.Columns("VoucherNo").Width = 1100 '2
        FrmTransacciones.DBGTransacciones.Columns("DescripcionMovimiento").Caption = "Descripcion" '3
        FrmTransacciones.DBGTransacciones.Columns("FacturaNo").Width = 1000 '4
        FrmTransacciones.DBGTransacciones.Columns("FacturaNo").Width = 1000 '4
        FrmTransacciones.DBGTransacciones.Columns("FacturaNo").Button = True '4
        FrmTransacciones.DBGTransacciones.Columns("ChequeNo").Width = 1000  '5
        FrmTransacciones.DBGTransacciones.Columns("ChequeNo").Caption = "Cheq/Rec" '5
        FrmTransacciones.DBGTransacciones.Columns("Clave").Button = True '6
        FrmTransacciones.DBGTransacciones.Columns("Clave").Locked = True '6
        FrmTransacciones.DBGTransacciones.Columns("Clave").Width = 800   '6
        FrmTransacciones.DBGTransacciones.Columns("TCambio").Caption = "Tasa Cambio"
        FrmTransacciones.DBGTransacciones.Columns("TCambio").Locked = True
        FrmTransacciones.DBGTransacciones.Columns("TCambio").NumberFormat = "##,##0.000000"
        FrmTransacciones.DBGTransacciones.Columns("TCambio").Width = 1200
        FrmTransacciones.DBGTransacciones.Columns("TCambio").Locked = True
        FrmTransacciones.DBGTransacciones.Columns("Debito").Width = 1200
        FrmTransacciones.DBGTransacciones.Columns("Debito").NumberFormat = "##,##0.00"
        FrmTransacciones.DBGTransacciones.Columns("Credito").Width = 1200
        FrmTransacciones.DBGTransacciones.Columns("Credito").NumberFormat = "##,##0.00"
        FrmTransacciones.DBGTransacciones.Columns("FechaTransaccion").Visible = False
        FrmTransacciones.DBGTransacciones.Columns("NPeriodo").Visible = False
        FrmTransacciones.DBGTransacciones.Columns("NTransaccion").Visible = False
        FrmTransacciones.DBGTransacciones.Columns("Fuente").Visible = False
        FrmTransacciones.DBGTransacciones.Columns("FechaTasas").Visible = False
        FrmTransacciones.DBGTransacciones.Columns("NumeroMovimiento").Visible = False
        FrmTransacciones.DBGTransacciones.Columns("Periodo").Visible = False
        FrmTransacciones.DBGTransacciones.Columns("FechaDescuento").Visible = False
        FrmTransacciones.DBGTransacciones.Columns("DescuentoDisponible").Visible = False
        FrmTransacciones.DBGTransacciones.Columns("FechaVence").Visible = False
        FrmTransacciones.DBGTransacciones.Columns("CodCuentaProveedor").Visible = False
        FrmTransacciones.DBGTransacciones.Columns("TipoFactura").Visible = False
        FrmTransacciones.DBGTransacciones.Columns("NTransaccion").Visible = False
     FrmTransacciones.TxtFecha.Enabled = False
     FrmTransacciones.TxtPeriodo.Enabled = False
     FrmTransacciones.TxtFuente.Enabled = False
     FrmTransacciones.TxtNTransacciones.Enabled = False
  End If
  End Select

  
  
  
 Unload Me
Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub CmdX_Click()
On Error GoTo TipoErrs
Unload Me
Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub DbgrProducto_DblClick()
On Error GoTo TipoErrs
CmdPegar.Value = True
Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub DbgrProducto_FilterChange()
  Me.DbgrProducto.PostMsg (418)

'On Error GoTo TipoErrs
'''Dim Filtro As String
'''Set cols = DbgrProducto.Columns
'''Dim c As Integer
''
'''c = DbgrProducto.col
'''DbgrProducto.HoldFields
'''Filtro = getFilter()
'''DtaProductos.Recordset.Filter = Filtro
'''DbgrProducto.col = c
'''DbgrProducto.EditActive = True
'
'    If Me.DbgrProducto.Columns(0).FilterText = "" And Me.DbgrProducto.Columns(1).FilterText = "" And Me.DbgrProducto.Columns(2).FilterText = "" Then
'     rs.Filter = ""
'     Exit Sub
'    End If
'
'
'    Dim col As TrueOleDBGrid80.Column
'    Dim cols As TrueOleDBGrid80.Columns
'    Dim Res As String
'
'
'    On Error Resume Next
'    Set cols = Me.DbgrProducto.Columns
'    Dim c As Integer
'
'
'    c = DbgrProducto.col
'    DbgrProducto.HoldFields
'    SQL = rs.Filter
'    rs.Filter = getFilter(col, cols)
'    If rs.EOF Then
'      MsgBox "No Existen Registros", vbInformation, "Zeus Contabilidad"
'      Res = LimpiarFilter(col, cols)
'      rs.Filter = ""
'    End If
'
'
''
''           Dim sb As String
''
''
''           For Each col In cols
''              If Len(col.FilterText) > 0 Then
''                 If Len(sb) > 0 Then
''                    sb = sb & (" AND ")
''                 End If
''
''                   sb = sb & ((col.DataField + " like " + "'" + col.FilterText + "*'"))
''
''               End If
''            Next col
''
''           rs.Filter = sb
'
'
'
'    DbgrProducto.col = c
'    DbgrProducto.EditActive = True
'
'Exit Sub
'TipoErrs:
' MsgBox err.Description
End Sub


Private Function getFilter(col As TrueOleDBGrid80.Column, cols As TrueOleDBGrid80.Columns) As String
'Creates the SQL statement in adodc1.recordset.filter
'and only filters text currently. It must be modified to
'filter other data types.
On Error GoTo TipoErrs

Dim tmp As String
Dim n As Integer
Dim x As Integer

For Each col In cols
    If Trim(col.FilterText) <> "" Then
        n = n + 1
        If n > 1 Then tmp = tmp & " AND "
        Select Case rs.Fields(x).Type
        Case adVarWChar, adVarChar: tmp = tmp & "[" & col.DataField & "] LIKE '%" & col.FilterText & "%'"
        Case adInteger, adNumeric: tmp = tmp & "[" & col.DataField & "] = " & col.FilterText
        Case adDBTimeStamp: tmp = tmp & "[" & col.DataField & "] = #" & col.FilterText & "#"
        End Select
    End If
    x = x + 1
Next col

If tmp <> "" Then
  getFilter = tmp
End If


Exit Function
TipoErrs:
 MsgBox err.Description

End Function

Private Function LimpiarFilter(col As TrueOleDBGrid80.Column, cols As TrueOleDBGrid80.Columns) As String
'Creates the SQL statement in adodc1.recordset.filter
'and only filters text currently. It must be modified to
'filter other data types.
On Error GoTo TipoErrs

Dim tmp As String
Dim n As Integer
Dim x As Integer

For Each col In cols
    col.FilterText = ""

    
    x = x + 1
Next col


Exit Function
TipoErrs:
 MsgBox err.Description

End Function





Private Sub DbgrProducto_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo TipoErrs
' Dim consulta
' Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
' If Shift = 1 Then
'  If Not KeyCode = 16 Then
'   Select Case KeyCode
'    Case 57
'     Lectura = "("
'    Case 219
'     Lectura = "-"
'    Case 48
'     Lectura = ")"
'   End Select
'  Else
'   Exit Sub
'  End If
' Else
'  'Si no Preciona Shift leo la tecla
'    Lectura = KeyCode
' '////////Chequeo si utiliza las Flechas direccioneales//////////
'   If Lectura = 40 Then
'    Exit Sub
'   End If
'   If Lectura = 38 Then
'    Exit Sub
'   End If
'   If Lectura = 39 Then
'    Exit Sub
'   End If
'   If Lectura = 37 Then
'    Exit Sub
'   End If
' '///////////Fin de la Busqueda Direccional////////////////////////
'
'  If Lectura = 8 Then
'   If Not Respuesta = "" Then
'    Respuesta = Mid$(Respuesta, 1, Len(Respuesta) - 1)
'    Lectura = ""
'    If Not Origen = "" Then
'     Origen = Mid$(Origen, 1, Len(Origen) - 1)
'    End If
'   End If
'  Else
'
'   LeeTecla
'  End If
' End If
'
'
'
' If KeyCode = 13 Then
'  Me.CmdPegar.Value = True
' Else
'   Select Case QueProducto
'
'
'
'           Case "CuentaOriginal"
'         If Orden = False Then
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'
'
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'           End If
'              DbgrProducto.Columns(0).Caption = "Código Cuenta"
'              DbgrProducto.Columns(0).Width = 2000
'              DbgrProducto.Columns(1).Width = 4200
'              DbgrProducto.Columns(3).Width = 2000
'              Orden = False
'            'Si el orden  Cambia a codigo entra a esta opcion
'            Else
'              If Respuesta = "" Or Lectura = "%" Then
'                 Respuesta = Lectura
'
'                 DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'
'                 DtaProductos.Refresh
'                 If Lectura = "%" Then
'                  Respuesta = ""
'                  Lectura = ""
'                 End If
'              Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'              DtaProductos.Refresh
'           End If
'         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(1).Width = 2000
'              Orden = True
'
'          End If
'
'           Case "CuentaGastos"
'         If Orden = False Then
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'           End If
'              DbgrProducto.Columns(0).Caption = "Código Cuenta"
'              DbgrProducto.Columns(0).Width = 2000
'              DbgrProducto.Columns(1).Width = 4200
'              DbgrProducto.Columns(3).Width = 2000
'              Orden = False
'            'Si el orden  Cambia a codigo entra a esta opcion
'            Else
'              If Respuesta = "" Or Lectura = "%" Then
'                 Respuesta = Lectura
'                 DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'
'                 DtaProductos.Refresh
'                 If Lectura = "%" Then
'                  Respuesta = ""
'                  Lectura = ""
'                 End If
'              Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'              DtaProductos.Refresh
'           End If
'         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(1).Width = 2000
'              Orden = True
'
'          End If
'
'        Case "CuentaDepreciacion"
'         If Orden = False Then
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'           End If
'              DbgrProducto.Columns(0).Caption = "Código Cuenta"
'              DbgrProducto.Columns(0).Width = 2000
'              DbgrProducto.Columns(1).Width = 4200
'              DbgrProducto.Columns(3).Width = 2000
'              Orden = False
'            'Si el orden  Cambia a codigo entra a esta opcion
'            Else
'              If Respuesta = "" Or Lectura = "%" Then
'                 Respuesta = Lectura
'                 DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'
'                 DtaProductos.Refresh
'                 If Lectura = "%" Then
'                  Respuesta = ""
'                  Lectura = ""
'                 End If
'              Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'              DtaProductos.Refresh
'           End If
'         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(1).Width = 2000
'              Orden = True
'
'          End If
'
'
'        Case "Auxiliar"
'         If Orden = False Then
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'           End If
'              DbgrProducto.Columns(0).Caption = "Código Cuenta"
'              DbgrProducto.Columns(0).Width = 2000
'              DbgrProducto.Columns(1).Width = 4200
'              DbgrProducto.Columns(3).Width = 2000
'              Orden = False
'            'Si el orden  Cambia a codigo entra a esta opcion
'            Else
'              If Respuesta = "" Or Lectura = "%" Then
'                 Respuesta = Lectura
'                 DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'
'                 DtaProductos.Refresh
'                 If Lectura = "%" Then
'                  Respuesta = ""
'                  Lectura = ""
'                 End If
'              Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'              DtaProductos.Refresh
'           End If
'         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(1).Width = 2000
'              Orden = True
'
'          End If
'
'      Case "Periodo"
'         If Orden = False Then
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') AND ((Cuentas.TipoCuenta)='Capital')) ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') AND ((Cuentas.TipoCuenta)='Capital')) ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'           End If
'              DbgrProducto.Columns(0).Caption = "Código Cuenta"
'              DbgrProducto.Columns(1).Width = 4200
'              Orden = False
'            'Si el orden  Cambia a codigo entra a esta opcion
'            Else
'              If Respuesta = "" Or Lectura = "%" Then
'                 Respuesta = Lectura
'                 DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE     (Cuentas.TipoCuenta = 'Capital') AND (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'                 DtaProductos.Refresh
'                 If Lectura = "%" Then
'                  Respuesta = ""
'                  Lectura = ""
'                 End If
'              Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE     (Cuentas.TipoCuenta = 'Capital') AND (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'              DtaProductos.Refresh
'           End If
'         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(1).Width = 2000
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(0).Width = 4200
'              Orden = True
'
'          End If
'
'        Case "CuentaReportes2"
'         If Orden = False Then
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'
'
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'           End If
'              DbgrProducto.Columns(0).Caption = "Código Cuenta"
'              DbgrProducto.Columns(0).Width = 2000
'              DbgrProducto.Columns(1).Width = 4200
'              DbgrProducto.Columns(3).Width = 2000
'              Orden = False
'            'Si el orden  Cambia a codigo entra a esta opcion
'            Else
'              If Respuesta = "" Or Lectura = "%" Then
'                 Respuesta = Lectura
'                 DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'
'                 DtaProductos.Refresh
'                 If Lectura = "%" Then
'                  Respuesta = ""
'                  Lectura = ""
'                 End If
'              Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'              DtaProductos.Refresh
'           End If
'         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(1).Width = 2000
'              Orden = True
'
'          End If
'
'        Case "CuentaReportes"
'         If Orden = False Then
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'
'
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'           End If
'              DbgrProducto.Columns(0).Caption = "Código Cuenta"
'              DbgrProducto.Columns(0).Width = 2000
'              DbgrProducto.Columns(1).Width = 4200
'              DbgrProducto.Columns(3).Width = 2000
'              Orden = False
'            'Si el orden  Cambia a codigo entra a esta opcion
'            Else
'              If Respuesta = "" Or Lectura = "%" Then
'                 Respuesta = Lectura
'                 DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'
'                 DtaProductos.Refresh
'                 If Lectura = "%" Then
'                  Respuesta = ""
'                  Lectura = ""
'                 End If
'              Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'              DtaProductos.Refresh
'           End If
'         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(1).Width = 2000
'              Orden = True
'
'          End If
'
'
'        Case "Cuenta"
'         If Orden = False Then
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'
'
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'           End If
'              DbgrProducto.Columns(0).Caption = "Código Cuenta"
'              DbgrProducto.Columns(0).Width = 2000
'              DbgrProducto.Columns(1).Width = 4200
'              DbgrProducto.Columns(3).Width = 2000
'              Orden = False
'            'Si el orden  Cambia a codigo entra a esta opcion
'            Else
'              If Respuesta = "" Or Lectura = "%" Then
'                 Respuesta = Lectura
'                 DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'
'                 DtaProductos.Refresh
'                 If Lectura = "%" Then
'                  Respuesta = ""
'                  Lectura = ""
'                 End If
'              Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'              DtaProductos.Refresh
'           End If
'         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(1).Width = 2000
'              Orden = True
'
'          End If
'
'        Case "CuentaMayor"
'         If Orden = False Then
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'           End If
'              DbgrProducto.Columns(0).Caption = "Código Cuenta"
'              DbgrProducto.Columns(0).Width = 2000
'              DbgrProducto.Columns(1).Width = 4200
'              DbgrProducto.Columns(3).Width = 2000
'              Orden = False
'            'Si el orden  Cambia a codigo entra a esta opcion
'            Else
'              If Respuesta = "" Or Lectura = "%" Then
'                 Respuesta = Lectura
'                 DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'
'                 DtaProductos.Refresh
'                 If Lectura = "%" Then
'                  Respuesta = ""
'                  Lectura = ""
'                 End If
'              Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'              DtaProductos.Refresh
'           End If
'         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(1).Width = 2000
'              Orden = True
'
'          End If
'
'        Case "Transferencia2"
'         If Orden = False Then
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'           End If
'              DbgrProducto.Columns(0).Caption = "Código Cuenta"
'              DbgrProducto.Columns(0).Width = 2000
'              DbgrProducto.Columns(1).Width = 4200
'              DbgrProducto.Columns(3).Width = 2000
'              Orden = False
'            'Si el orden  Cambia a codigo entra a esta opcion
'            Else
'              If Respuesta = "" Or Lectura = "%" Then
'                 Respuesta = Lectura
'                 DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'
'                 DtaProductos.Refresh
'                 If Lectura = "%" Then
'                  Respuesta = ""
'                  Lectura = ""
'                 End If
'              Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'              DtaProductos.Refresh
'           End If
'         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(1).Width = 2000
'              Orden = True
'
'          End If
'
'        Case "Transferencia1"
'         If Orden = False Then
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'           End If
'              DbgrProducto.Columns(0).Caption = "Código Cuenta"
'              DbgrProducto.Columns(0).Width = 2000
'              DbgrProducto.Columns(1).Width = 4200
'              DbgrProducto.Columns(3).Width = 2000
'              Orden = False
'            'Si el orden  Cambia a codigo entra a esta opcion
'            Else
'              If Respuesta = "" Or Lectura = "%" Then
'                 Respuesta = Lectura
'                 DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'
'                 DtaProductos.Refresh
'                 If Lectura = "%" Then
'                  Respuesta = ""
'                  Lectura = ""
'                 End If
'              Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'              DtaProductos.Refresh
'           End If
'         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(1).Width = 2000
'              Orden = True
'
'          End If
'
'        Case "ContratistaCheque"
'         If Orden = False Then
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'
'              DtaProductos.RecordSource = "SELECT Contactos.CodigoCuenta,  Contactos.Beneficiario,Contactos.Ciudad, Contactos.Telefono From Contactos Where (((Contactos.CodigoCuenta) Like '" & Respuesta & "%')) ORDER BY Contactos.CodigoCuenta"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Contactos.CodigoCuenta, Contactos.Beneficiario, Contactos.Ciudad, Contactos.Telefono From Contactos Where (((Contactos.CodigoCuenta) Like '" & Respuesta & "%')) ORDER BY Contactos.CodigoCuenta"
'              DtaProductos.Refresh
'           End If
'              DbgrProducto.Columns(0).Caption = "Código Cuenta"
'              DbgrProducto.Columns(0).Width = 2000
'              DbgrProducto.Columns(1).Width = 4200
'              DbgrProducto.Columns(3).Width = 2000
'              Orden = False
'            'Si el orden  Cambia a codigo entra a esta opcion
'            Else
'              If Respuesta = "" Or Lectura = "%" Then
'                 Respuesta = Lectura
'                 DtaProductos.RecordSource = "SELECT Contactos.Beneficiario, Contactos.CodigoCuenta, Contactos.Ciudad, Contactos.Telefono From Contactos Where (((Contactos.Beneficiario) Like '" & Respuesta & "%')) ORDER BY Contactos.Beneficiario"
'
'                 DtaProductos.Refresh
'                 If Lectura = "%" Then
'                  Respuesta = ""
'                  Lectura = ""
'                 End If
'              Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Contactos.Beneficiario, Contactos.CodigoCuenta, Contactos.Ciudad, Contactos.Telefono From Contactos Where (((Contactos.Benefiario) Like '" & Respuesta & "%')) ORDER BY Contactos.Beneficiario"
'              DtaProductos.Refresh
'           End If
'         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(1).Width = 2000
'              Orden = True
'
'          End If
'
'        Case "Contratista"
'         If Orden = False Then
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'
'              DtaProductos.RecordSource = "SELECT Contactos.CodigoCuenta, Contactos.Beneficiario, Contactos.Ciudad, Contactos.Telefono From Contactos Where (((Contactos.CodigoCuenta) Like '" & Respuesta & "%')) ORDER BY Contactos.CodigoCuenta"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Contactos.CodigoCuenta,Contactos.Beneficiario, Contactos.Ciudad, Contactos.Telefono From Contactos Where (((Contactos.CodigoCuenta) Like '" & Respuesta & "%')) ORDER BY Contactos.CodigoCuenta"
'              DtaProductos.Refresh
'           End If
'              DbgrProducto.Columns(0).Caption = "Código Cuenta"
'              DbgrProducto.Columns(1).Width = 4200
'              DbgrProducto.Columns(0).Width = 2000
'              DbgrProducto.Columns(3).Width = 2000
'              Orden = False
'            'Si el orden  Cambia a codigo entra a esta opcion
'            Else
'              If Respuesta = "" Or Lectura = "%" Then
'                 Respuesta = Lectura
'                 DtaProductos.RecordSource = "SELECT Contactos.Beneficiario, Contactos.CodigoCuenta,Contactos.Ciudad, Contactos.Telefono From Contactos Where (((Contactos.Beneficiario) Like '" & Respuesta & "%')) ORDER BY Contactos.Beneficiario"
'
'                 DtaProductos.Refresh
'                 If Lectura = "%" Then
'                  Respuesta = ""
'                  Lectura = ""
'                 End If
'              Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Contactos.Beneficiario, Contactos.CodigoCuenta,Contactos.Ciudad, Contactos.Telefono From Contactos Where (((Contactos.Beneficiario) Like '" & Respuesta & "%')) ORDER BY Contactos.Beneficiario"
'              DtaProductos.Refresh
'           End If
'         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(1).Width = 2000
'              Orden = True
'
'          End If
'
'
'
'     Case "Cheque"
'         If Orden = False Then
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'           End If
'              DbgrProducto.Columns(0).Caption = "Código Cuenta"
'              DbgrProducto.Columns(0).Width = 2000
'              DbgrProducto.Columns(1).Width = 4200
'              DbgrProducto.Columns(3).Width = 2000
'              Orden = False
'            'Si el orden  Cambia a codigo entra a esta opcion
'            Else
'              If Respuesta = "" Or Lectura = "%" Then
'                 Respuesta = Lectura
'                 DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas,Cuentas.CodCuentas, Cuentas.TipoCuenta, GrupoCuentas.DescripcionGrupo FROM GrupoCuentas INNER JOIN Cuentas ON GrupoCuentas.CodGrupo = Cuentas.CodGrupo Where (((Cuentas.DescripcionCuentas) Like '" & Respuesta & "%')) ORDER BY Cuentas.DescripcionCuentas"
'                 DtaProductos.Refresh
'                 If Lectura = "%" Then
'                  Respuesta = ""
'                  Lectura = ""
'                 End If
'              Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, GrupoCuentas.DescripcionGrupo FROM GrupoCuentas INNER JOIN Cuentas ON GrupoCuentas.CodGrupo = Cuentas.CodGrupo Where (((Cuentas.DescripcionCuentas) Like '" & Respuesta & "%')) ORDER BY Cuentas.DescripcionCuentas"
'              DtaProductos.Refresh
'           End If
'         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(1).Width = 2000
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(0).Width = 4200
'              Orden = True
'
'          End If
'         Case "AuxiliarTransacciones"
'         If Orden = False Then
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'           End If
'              DbgrProducto.Columns(0).Caption = "Código Cuenta"
'              DbgrProducto.Columns(1).Width = 4200
'              Orden = False
'            'Si el orden  Cambia a codigo entra a esta opcion
'            Else
'              If Respuesta = "" Or Lectura = "%" Then
'                 Respuesta = Lectura
'                 DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'                 DtaProductos.Refresh
'                 If Lectura = "%" Then
'                  Respuesta = ""
'                  Lectura = ""
'                 End If
'              Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'              DtaProductos.Refresh
'           End If
' Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(1).Width = 2000
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(0).Width = 4200
'              Orden = True
'
'          End If
'
'      Case "Transacciones"
'         If Orden = False Then
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'           End If
'              DbgrProducto.Columns(0).Caption = "Código Cuenta"
'              DbgrProducto.Columns(1).Width = 4200
'              Orden = False
'            'Si el orden  Cambia a codigo entra a esta opcion
'            Else
'              If Respuesta = "" Or Lectura = "%" Then
'                 Respuesta = Lectura
'                 DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'                 DtaProductos.Refresh
'                 If Lectura = "%" Then
'                  Respuesta = ""
'                  Lectura = ""
'                 End If
'              Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'              DtaProductos.Refresh
'           End If
' Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(1).Width = 2000
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(0).Width = 4200
'              Orden = True
'
'          End If
'
'    Case "NTransacciones"
'
'         Mes = Month(FrmTransacciones.TxtFecha.Value)
'         Año = Year(FrmTransacciones.TxtFecha.Value)
'         FechaIni = CDate("1/" & Month(FrmTransacciones.TxtFecha.Value) & "/" & Year(FrmTransacciones.TxtFecha.Value))
'         FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
'         NumFecha1 = FechaIni
'         NumFecha2 = FechaFin
'
'
'
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'              DtaProductos.RecordSource = "SELECT IndiceTransaccion.NumeroMovimiento,IndiceTransaccion.FechaTransaccion, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "') AND ((IndiceTransaccion.NumeroMovimiento) Like '" & Respuesta & "%')) ORDER BY IndiceTransaccion.NumeroMovimiento"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.FechaTransaccion, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "') AND ((IndiceTransaccion.NumeroMovimiento) Like '" & Respuesta & "%')) ORDER BY IndiceTransaccion.NumeroMovimiento"
'              DtaProductos.Refresh
'           End If
'              DbgrProducto.Columns(1).Caption = "No.Transaccion"
'
'
'         Case "ContratistasM"
'         If Orden = False Then
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'           End If
'              DbgrProducto.Columns(0).Caption = "Código Cuenta"
'              DbgrProducto.Columns(1).Width = 4200
'              Orden = False
'            'Si el orden  Cambia a codigo entra a esta opcion
'            Else
'              If Respuesta = "" Or Lectura = "%" Then
'                 Respuesta = Lectura
'                 DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'                 DtaProductos.Refresh
'                 If Lectura = "%" Then
'                  Respuesta = ""
'                  Lectura = ""
'                 End If
'              Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'              DtaProductos.Refresh
'           End If
' Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(1).Width = 2000
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(0).Width = 4200
'              Orden = True
'
'          End If
'
'             Case "MisCode"
'         If Orden = False Then
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.CodCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.CodCuentas"
'              DtaProductos.Refresh
'           End If
'              DbgrProducto.Columns(0).Caption = "Código Cuenta"
'              DbgrProducto.Columns(1).Width = 4200
'              Orden = False
'            'Si el orden  Cambia a codigo entra a esta opcion
'            Else
'              If Respuesta = "" Or Lectura = "%" Then
'                 Respuesta = Lectura
'                 DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'                 DtaProductos.Refresh
'                 If Lectura = "%" Then
'                  Respuesta = ""
'                  Lectura = ""
'                 End If
'              Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY Cuentas.DescripcionCuentas"
'              DtaProductos.Refresh
'           End If
' Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(1).Width = 2000
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(0).Width = 4200
'              Orden = True
'
'          End If
'
'
'
'
'    End Select
'    'Si no lo encuentra el producto Limpio la pantalla
'    If DtaProductos.Recordset.EOF Then
'
'       Select Case QueProducto
'     Case "CuentaOriginal"
'         SqlConsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = SqlConsulta
'         DtaProductos.Refresh
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'        Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(1).Width = 2000
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'
'     Case "CuentaGastos"
'         SqlConsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = SqlConsulta
'         DtaProductos.Refresh
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'        Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(1).Width = 2000
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'
'     Case "CuentaDepreciacion"
'         SqlConsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = SqlConsulta
'         DtaProductos.Refresh
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'        Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(1).Width = 2000
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'
'     Case "Auxiliar"
'         SqlConsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = SqlConsulta
'         DtaProductos.Refresh
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'        Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(1).Width = 2000
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'
'     Case "Periodo"
'         SqlConsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta,Cuentas.TipoMoneda FROM GrupoCuentas INNER JOIN Cuentas ON GrupoCuentas.CodGrupo = Cuentas.CodGrupo Where (((Cuentas.TipoCuenta) = 'Capital')) ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = SqlConsulta
'         DtaProductos.Refresh
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'        Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(1).Width = 2000
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'     Case "Cuenta"
'         SqlConsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = SqlConsulta
'         DtaProductos.Refresh
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'        Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(1).Width = 2000
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'
'     Case "CuentaMayor"
'         SqlConsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = SqlConsulta
'         DtaProductos.Refresh
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'        Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(1).Width = 2000
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'
'     Case "Transferencia2"
'         SqlConsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = SqlConsulta
'         DtaProductos.Refresh
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'        Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(1).Width = 2000
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'
'     Case "Transferencia1"
'         SqlConsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = SqlConsulta
'         DtaProductos.Refresh
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'        Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(1).Width = 2000
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'     Case "Contratista"
'         SqlConsulta = "SELECT Contactos.Beneficiario, Contactos.CodigoCuenta, Contactos.Ciudad, Contactos.Telefono From Contactos ORDER BY Contactos.Beneficiario"
'         DtaProductos.RecordSource = SqlConsulta
'                  DtaProductos.Refresh
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'
'         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(1).Width = 2000
'         'DbgrProducto.Columns(1).Caption = "Código Cuenta"
'         'DbgrProducto.Columns(0).Width = 4200
'
'         Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
'
'
'     Case "ContratistaCheque"
'         SqlConsulta = "SELECT Contactos.Beneficiario, Contactos.CodigoCuenta , Contactos.Ciudad, Contactos.Telefono From Contactos ORDER BY Contactos.Beneficiario"
'         DtaProductos.RecordSource = SqlConsulta
'                  DtaProductos.Refresh
'         Respuesta = ""
'         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(0).Caption = "Código Cuentas"
'              DbgrProducto.Columns(1).Width = 2000
'         'DbgrProducto.Columns(1).Caption = "Código Cuenta"
'         'DbgrProducto.Columns(0).Width = 4200
'
'         Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
'
'        Lectura = ""
'        Origen = ""
'      Case "ContratistasM"
'         SqlConsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = SqlConsulta
'         DtaProductos.Refresh
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'        Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(1).Width = 2000
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'
'        Case "MisCode"
'         SqlConsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = SqlConsulta
'         DtaProductos.Refresh
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'        Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(1).Width = 2000
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'     Case "AuxiliarTransacciones"
'         SqlConsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = SqlConsulta
'         DtaProductos.Refresh
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'        Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(1).Width = 2000
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'
'     Case "Transacciones"
'         SqlConsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = SqlConsulta
'         DtaProductos.Refresh
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'        Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(1).Width = 2000
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'     Case "NTransacciones"
'         FrmTransacciones.DBGTransacciones.Enabled = True
'         Mes = Month(FrmTransacciones.TxtFecha.Value)
'         Año = Year(FrmTransacciones.TxtFecha.Value)
'         FechaIni = CDate("1/" & Month(FrmTransacciones.TxtFecha.Value) & "/" & Year(FrmTransacciones.TxtFecha.Value))
'         FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
'         NumFecha1 = FechaIni
'         NumFecha2 = FechaFin
'         SqlConsulta = "SELECT IndiceTransaccion.NumeroMovimiento,IndiceTransaccion.FechaTransaccion, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & " ))"
'         DtaProductos.RecordSource = SqlConsulta
'         DtaProductos.Refresh
'         Me.CmdOrden.Enabled = False
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
'     Case "Cheque"
'         SqlConsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = SqlConsulta
'         DtaProductos.Refresh
'         Respuesta = ""
'         Origen = ""
'         Lectura = ""
' Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
'         Me.DbgrProducto.Columns(1).Width = 2000
'         Me.DbgrProducto.Columns(3).Width = 2000
'              DbgrProducto.Columns(1).Caption = "Código Cuentas"
'              DbgrProducto.Columns(0).Width = 4200
'         Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
'
'
'
'     End Select
'     Me.Caption = "Buscar:"
'    Else
'    Origen = Origen & Lectura
'    If Lectura = " " Then
'      Origen = Mid$(Origen, 1, Len(Origen) - 1)
'      Origen = Origen & "_"
'    End If
'    Me.Caption = "Buscar: " & Origen
'    End If
'  End If
'
'
'Exit Sub
'TipoErrs:
'  ControlErrores
'  Unload Me




End Sub
Public Function CadenaNumeros(cadena As String) As Double
  Dim Index As Integer, caracteres As Integer, numeros As String, Codigo As String

    Index = 1
    caracteres = Len(cadena)
    While Not Index = caracteres
        numeros = Mid(cadena, Index, 1)
            If IsNumeric(numeros) Then
                Codigo = Codigo & numeros
            End If
            If IsNumeric(numeros) = False Then
'                Label5.Caption = Label5.Caption & numeros
            End If
        Index = Index + 1
    Wend
    
'    Label6.Caption = Len(cadena)
  
  If Codigo <> "" Then
    CadenaNumeros = Codigo
    
  End If


End Function





Private Sub Form_Load()
Dim mes As Double, Año As Double
Dim Fecha1 As String, Fecha2 As String
Dim Filtro As String
Dim c As Integer
Dim col As TrueOleDBGrid80.Column
Dim cols As TrueOleDBGrid80.Columns


 
On Error GoTo TipoErrs
MDIPrimero.Skin1.ApplySkin hWnd
 Me.DbgrProducto.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DbgrProducto.OddRowStyle.BackColor = &H80000005
 Me.DbgrProducto.AlternatingRowStyle = True
 'AZUL Y BLANCO COMO LA PATRIA
 
If cnx.State = adStateClosed Then
    cnx.ConnectionString = Conexion
    cnx.Open
End If


With Me.DtaTasas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaProductos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaConsulta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.AdoBusca
   .ConnectionString = Conexion
End With

Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas ORDER BY Tasas.FechaTasas"
Me.DtaTasas.Refresh


Origen = ""
Dim sqlconsulta As String
 'CmdX.Visible = True
Orden = True
Select Case QueProducto

      Case "Presupuesto"
          sqlconsulta = "SELECT KeyGrupo, DescripcionGrupo From EstructuraPresupuesto"
        
                With rs
                  .CursorLocation = adUseClient
                  .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
                End With
                
                Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Pres."
         Me.DbgrProducto.Columns(0).Width = 1000
         DbgrProducto.Columns(1).Width = 3000

      Case "CuentaContable"
          sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"

        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         DbgrProducto.Columns(0).Width = 4200
         


       Case "Fuente"
         sqlconsulta = "SELECT DISTINCT Fuente From IndiceTransaccion ORDER BY Fuente"
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh

        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        
        Me.DbgrProducto.DataSource = rs

         Respuesta = ""
         Me.DbgrProducto.Columns(0).Caption = "Fuente"
         Me.DbgrProducto.Columns(0).Width = 2000




       Case "Departamento"
         sqlconsulta = "SELECT  * From GrupoCuentas"
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh

        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs

         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Nombre Dpto"
         Me.DbgrProducto.Columns(0).Caption = "Codigo Dpto"
         Me.DbgrProducto.Columns(0).Width = 2000
         DbgrProducto.Columns(1).Width = 4200
       
       Case "Prorrateo2"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo " & _
                       "FROM  Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo " & _
                       "WHERE (Cuentas.TipoCuenta = N'Costos') OR (Cuentas.TipoCuenta = N'Gastos') ORDER BY Cuentas.DescripcionCuentas "
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh

        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs

         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         DbgrProducto.Columns(0).Width = 4200

       Case "Prorrateo"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo " & _
                       "FROM  Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo " & _
                       "WHERE (Cuentas.TipoCuenta = N'Costos') OR (Cuentas.TipoCuenta = N'Gastos') ORDER BY Cuentas.DescripcionCuentas "
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         DbgrProducto.Columns(0).Width = 4200
       
       Case "CuentaOriginal"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         DbgrProducto.Columns(0).Width = 4200
       
       
       Case "CuentaGastos"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
         'SqlConsulta = "SELECT Cuentas.DescripcionCuentas,Cuentas.CodCuentas,Cuentas.TipoCuenta, GrupoCuentas.DescripcionGrupo FROM GrupoCuentas INNER JOIN Cuentas ON GrupoCuentas.CodGrupo = Cuentas.CodGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         DbgrProducto.Columns(0).Width = 4200

       
       Case "CuentaDepreciacion"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
         'SqlConsulta = "SELECT Cuentas.DescripcionCuentas,Cuentas.CodCuentas,Cuentas.TipoCuenta, GrupoCuentas.DescripcionGrupo FROM GrupoCuentas INNER JOIN Cuentas ON GrupoCuentas.CodGrupo = Cuentas.CodGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         DbgrProducto.Columns(0).Width = 4200
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs


       Case "Auxiliar"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
         'SqlConsulta = "SELECT Cuentas.DescripcionCuentas,Cuentas.CodCuentas,Cuentas.TipoCuenta, GrupoCuentas.DescripcionGrupo FROM GrupoCuentas INNER JOIN Cuentas ON GrupoCuentas.CodGrupo = Cuentas.CodGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh

        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         DbgrProducto.Columns(0).Width = 4200
         
         If FrmAuxiliarCuentas.DBCliente.Text <> "" Then
            Me.DbgrProducto.Columns(1).FilterText = FrmAuxiliarCuentas.DBCliente.Text
'            Set cols = DbgrProducto.Columns
'
'            c = DbgrProducto.col
'            DbgrProducto.HoldFields
'            Filtro = getFilter()
'            DtaProductos.Recordset.Filter = Filtro
'            DbgrProducto.col = c
'            DbgrProducto.EditActive = True
           
            'On Error GoTo errHandler
            On Error Resume Next
            Set cols = Me.DbgrProducto.Columns
'            Dim c As Integer
            
            c = DbgrProducto.col
            DbgrProducto.HoldFields
            Sql = rs.Filter
            rs.Filter = getFilter(col, cols)
            DbgrProducto.col = c
            DbgrProducto.EditActive = True
         End If


     Case "Periodo"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta,Cuentas.TipoMoneda, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE (Cuentas.TipoCuenta = 'Capital') ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Origen = ""
         Lectura = ""
        Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
         Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
         Respuesta = ""
         Origen = ""
         Lectura = ""
       Case "Cuenta"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         DbgrProducto.Columns(0).Width = 4200
         
         If FrmCuentas.DBCliente.Text <> "" Then
            Me.DbgrProducto.Columns(1).FilterText = FrmCuentas.DBCliente.Text
            On Error Resume Next
            Set cols = Me.DbgrProducto.Columns
'            Dim c As Integer
            
            c = DbgrProducto.col
            DbgrProducto.HoldFields
            Sql = rs.Filter
            rs.Filter = getFilter(col, cols)
            DbgrProducto.col = c
            DbgrProducto.EditActive = True
'            Set cols = DbgrProducto.Columns
'
'            c = DbgrProducto.col
'            DbgrProducto.HoldFields
'            Filtro = getFilter()
'            DtaProductos.Recordset.Filter = Filtro
'            DbgrProducto.col = c
'            DbgrProducto.EditActive = True
         End If
         
       Case "CuentaFactura"
       
         If FrmTransacciones.OptFacturaCompra.Value = True Then
           sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE  (Cuentas.TipoCuenta = 'Cuentas x Pagar') ORDER BY Cuentas.DescripcionCuentas"
         Else
           sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE  (Cuentas.TipoCuenta = 'Cuentas x Cobrar') ORDER BY Cuentas.DescripcionCuentas"
         End If
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         DbgrProducto.Columns(0).Width = 4200
         
         If FrmTransacciones.TDBProveedor.Text <> "" Then
            Me.DbgrProducto.Columns(1).FilterText = FrmTransacciones.TDBProveedor.Text
            On Error Resume Next
            Set cols = Me.DbgrProducto.Columns
'            Dim c As Integer
            
            c = DbgrProducto.col
            DbgrProducto.HoldFields
            Sql = rs.Filter
            rs.Filter = getFilter(col, cols)
            DbgrProducto.col = c
            DbgrProducto.EditActive = True
            
'            Set cols = DbgrProducto.Columns
'
'            c = DbgrProducto.col
'            DbgrProducto.HoldFields
'            Filtro = getFilter()
'            DtaProductos.Recordset.Filter = Filtro
'            DbgrProducto.col = c
'            DbgrProducto.EditActive = True
         End If
         
       Case "CuentaActivoFijo"
       

          sqlconsulta = "SELECT DescripcionActivo, CodCuenta, NumeroSerie, Marca From ActivoFijo"
 
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         DbgrProducto.Columns(0).Width = 4200
         
         If FrmActivoFijo.DBCodigo.Text <> "" Then
            Me.DbgrProducto.Columns(1).FilterText = FrmActivoFijo.DBCodigo.Text
            On Error Resume Next
            Set cols = Me.DbgrProducto.Columns
'            Dim c As Integer
            
            c = DbgrProducto.col
            DbgrProducto.HoldFields
            Sql = rs.Filter
            rs.Filter = getFilter(col, cols)
            DbgrProducto.col = c
            DbgrProducto.EditActive = True
'            Set cols = DbgrProducto.Columns
'
'            c = DbgrProducto.col
'            DbgrProducto.HoldFields
'            Filtro = getFilter()
'            DtaProductos.Recordset.Filter = Filtro
'            DbgrProducto.col = c
'            DbgrProducto.EditActive = True
         End If

       Case "CuentaReportes"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         DbgrProducto.Columns(0).Width = 4200
         
       Case "CuentaReportes2"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         DbgrProducto.Columns(0).Width = 4200


       Case "CuentaMayor"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         'Me.DbgrProducto.Columns(3).Width = 2000
          '    DbgrProducto.Columns(0).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
         'DbgrProducto.Columns(1).Caption = "Código Cuenta"
         'DbgrProducto.Columns(0).Width = 4200

       Case "Transferencia2"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         'Me.DbgrProducto.Columns(3).Width = 2000
          '    DbgrProducto.Columns(0).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
         'DbgrProducto.Columns(1).Caption = "Código Cuenta"
         'DbgrProducto.Columns(0).Width = 4200

       Case "Transferencia1"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         'Me.DbgrProducto.Columns(3).Width = 2000
          '    DbgrProducto.Columns(0).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
         'DbgrProducto.Columns(1).Caption = "Código Cuenta"
         'DbgrProducto.Columns(0).Width = 4200

       Case "MisCode"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         'Me.DbgrProducto.Columns(3).Width = 2000
          '    DbgrProducto.Columns(0).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
         'DbgrProducto.Columns(1).Caption = "Código Cuenta"
         'DbgrProducto.Columns(0).Width = 4200



        Case "ContratistasM"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         'Me.DbgrProducto.Columns(3).Width = 2000
          '    DbgrProducto.Columns(0).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
         'DbgrProducto.Columns(1).Caption = "Código Cuenta"
         'DbgrProducto.Columns(0).Width = 4200
    

     Case "Contratista"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = Sqlconsulta
'                  DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 4200
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(1).Width = 2000
         'DbgrProducto.Columns(1).Caption = "Código Cuenta"
         'DbgrProducto.Columns(0).Width = 4200


     Case "ContratistaCheque"
         sqlconsulta = "SELECT Contactos.Beneficiario, Contactos.CodigoCuenta,Contactos.Ciudad, Contactos.Telefono From Contactos ORDER BY Contactos.Beneficiario"
'         DtaProductos.RecordSource = Sqlconsulta
'                  DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 4200
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(1).Width = 2000
         'DbgrProducto.Columns(1).Caption = "Código Cuenta"
         'DbgrProducto.Columns(0).Width = 4200
   Case "AuxiliarTransacciones"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         'Me.DbgrProducto.Columns(3).Width = 2000
          '    DbgrProducto.Columns(0).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
         'DbgrProducto.Columns(1).Caption = "Código Cuenta"
         'DbgrProducto.Columns(0).Width = 4200

   Case "Transacciones"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(0).Width = 2000
         DbgrProducto.Columns(0).Width = 4200
         
        If FrmTransacciones.DBGTransacciones.Columns(0).Text <> "" Then
            Me.DbgrProducto.Columns(1).FilterText = FrmTransacciones.DBGTransacciones.Columns(0).Text
            On Error Resume Next
            Set cols = Me.DbgrProducto.Columns
'            Dim c As Integer
            
            c = DbgrProducto.col
            DbgrProducto.HoldFields
            Sql = rs.Filter
            rs.Filter = getFilter(col, cols)
            DbgrProducto.col = c
            DbgrProducto.EditActive = True
'            Set cols = DbgrProducto.Columns
'
'            c = DbgrProducto.col
'            DbgrProducto.HoldFields
'            Filtro = getFilter()
'            DtaProductos.Recordset.Filter = Filtro
'            DbgrProducto.col = c
'            DbgrProducto.EditActive = True
         End If

   Case "NTransacciones"
         FrmTransacciones.DBGTransacciones.Enabled = True
         mes = Month(FrmTransacciones.TxtFecha.Value)
         Año = Year(FrmTransacciones.TxtFecha.Value)
         FechaIni = CDate("1/" & Month(FrmTransacciones.TxtFecha.Value) & "/" & Year(FrmTransacciones.TxtFecha.Value))
         FechaFin = DateSerial(Año, mes + 1, 1 - 1)
         NumFecha1 = FechaIni
         NumFecha2 = FechaFin
         Fecha1 = Format(FechaIni, "yyyy-mm-dd")
         Fecha2 = Format(FechaFin, "yyyy-mm-dd")
'         SqlConsulta = "SELECT IndiceTransaccion.NumeroMovimiento,IndiceTransaccion.FechaTransaccion, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & " ))ORDER BY IndiceTransaccion.NumeroMovimiento"
'lo quite porque el between no agarra los extremos, es decir para febrero no me toma el primer dia el 1ero de febrero ni el 28
         sqlconsulta = "SELECT NumeroMovimiento, FechaTransaccion, DescripcionMovimiento, Fuente From IndiceTransaccion WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) ORDER BY NumeroMovimiento"
'         SqlConsulta = "SELECT IndiceTransaccion.NumeroMovimiento,IndiceTransaccion.FechaTransaccion, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Fuente From IndiceTransaccion WHERE IndiceTransaccion.FechaTransaccion>='" & Format(NumFecha1, "yyyymmdd") & "' And indicetransaccion.fechatransaccion<='" & Format(NumFecha2, "yyyymmdd") & "' ORDER BY IndiceTransaccion.NumeroMovimiento"
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
         
         Me.CmdOrden.Enabled = False
         Respuesta = ""
                   
                   DbgrProducto.Columns(2).Width = 4200
                   
    Case "Egreso"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
   Case "SolicitudPagos"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
              
     Case "SolicitudCheques"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo ORDER BY Cuentas.DescripcionCuentas"
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
              
    Case "Cheque"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo  ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200
              
    Case "ChequeBanco"
         sqlconsulta = "SELECT Cuentas.DescripcionCuentas, Cuentas.CodCuentas, Cuentas.TipoCuenta, Grupos.DescripcionGrupo FROM Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo WHERE Cuentas.TipoCuenta = 'Bancos' ORDER BY Cuentas.DescripcionCuentas"
'         DtaProductos.RecordSource = Sqlconsulta
'         DtaProductos.Refresh
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        Me.DbgrProducto.DataSource = rs
        
         Respuesta = ""
         Me.DbgrProducto.Columns(1).Caption = "Codigo Cuenta"
         Me.DbgrProducto.Columns(1).Width = 2000
         Me.DbgrProducto.Columns(3).Width = 2000
              DbgrProducto.Columns(1).Caption = "Código Cuentas"
              DbgrProducto.Columns(0).Width = 4200

    End Select
'   Me.DbgrProducto.MarqueeStyle = dbgHighlightCell

Exit Sub
TipoErrs:
MsgBox err.Description
Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo TipoErrs
Campo = False
Exit Sub
TipoErrs:
MsgBox err.Description
End Sub



Private Sub PegarSolicitudPago()

    FrmSolicitudPagos.DtaCuentas.Refresh
    FrmSolicitudPagos.DBGTransacciones.Columns(0).Text = rs("CodCuentas")
     Criterio = "CodCuentas='" & FrmSolicitudPagos.DBGTransacciones.Columns(0).Text & "'"
      FrmSolicitudPagos.DtaCuentas.Recordset.Find (Criterio)
       If Not FrmSolicitudPagos.DtaCuentas.Recordset.EOF Then
         mes = Month(FrmSolicitudPagos.TxtFecha.Value)
         Año = Year(FrmSolicitudPagos.TxtFecha.Value)
         FechaIni = CDate("1/" & Month(FrmSolicitudPagos.TxtFecha.Value) & "/" & Year(FrmSolicitudPagos.TxtFecha.Value))
         FechaFin = DateSerial(Año, mes + 1, 1 - 1)
         NumFecha1 = CDate(FechaIni)
         NumFecha2 = CDate(FechaFin)
 
        FrmSolicitudPagos.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
        FrmSolicitudPagos.DtaConsulta.Refresh
         If Not FrmSolicitudPagos.DtaConsulta.Recordset.EOF Then
          FrmSolicitudPagos.TxtPeriodo.Text = FrmSolicitudPagos.DtaConsulta.Recordset("Periodo")
            NumeroPeriodo = FrmSolicitudPagos.DtaConsulta.Recordset("NPeriodo")
            If Val(FrmSolicitudPagos.TxtNTransacciones.Text) = 0 Then
                NumeroTransaccion = FrmSolicitudPagos.DtaConsulta.Recordset("NTransacciones")
            Else
                NumeroTransaccion = FrmSolicitudPagos.TxtNTransacciones.Text
            End If
            EstadoPeriodo = FrmSolicitudPagos.DtaConsulta.Recordset("EstadoPeriodo")
      
        '////////////Edito los datos del Periodo///////////
         If Val(FrmSolicitudPagos.TxtNTransacciones.Text) = 0 Then
          
          
'         FrmSolicitudPagos.'DtaConsulta.Recordset.Edit
         FrmSolicitudPagos.DtaConsulta.Recordset("NTransacciones") = FrmSolicitudPagos.DtaConsulta.Recordset("NTransacciones") + 1
         FrmSolicitudPagos.DtaConsulta.Recordset.Update
          NumeroTransaccion = FrmSolicitudPagos.DtaConsulta.Recordset("NTransacciones")
         FrmSolicitudPagos.TxtNTransacciones.Text = NumeroTransaccion
          '////////Edito los Datos de los indices de Transacciones//////
         
         FrmSolicitudPagos.DtaIndice.Recordset.AddNew
         FrmSolicitudPagos.DtaIndice.Recordset("FechaTransaccion") = Format(FrmSolicitudPagos.TxtFecha.Value, "dd/mm/yyyy")
         FrmSolicitudPagos.DtaIndice.Recordset("NumeroMovimiento") = NumeroTransaccion
         FrmSolicitudPagos.DtaIndice.Recordset("Fuente") = FrmSolicitudPagos.TxtFuente.Text
         FrmSolicitudPagos.DtaIndice.Recordset("NPeriodo") = NumeroPeriodo
         
      Criterio = "CodCuentas='" & FrmSolicitudPagos.DBGTransacciones.Columns(0).Text & "'"
       FrmSolicitudPagos.DtaCuentas.Recordset.Find (Criterio)
       If Not FrmSolicitudPagos.DtaCuentas.Recordset.EOF Then
        TipoMoneda = FrmSolicitudPagos.DtaCuentas.Recordset("TipoMoneda")

         Select Case TipoMoneda
            Case "Córdobas"
            
                      Fecha = FrmSolicitudPagos.TxtFecha.Value
                      Fechas = Format(FrmSolicitudPagos.TxtFecha.Value, "yyyy/mm/dd")
'                      FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE (FechaTasas = '" & Format(FrmSolicitudPagos.TxtFecha.Value, "yyyymmdd") & "')"
                      FrmSolicitudPagos.DtaTasas.Refresh
                If Not FrmSolicitudPagos.DtaTasas.Recordset.EOF Then
                 MontoTasa = FrmSolicitudPagos.DtaTasas.Recordset("MontoCordobas")
                 If MontoTasa = 0 Then
                   MontoTasa = 1
                 End If
                 Select Case FrmSolicitudPagos.CmbMoneda.Text
                  Case "Córdobas"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
                  Case "Dólares"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Libras"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = MontoTasa
                   ' FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                    
                 End Select
                Else
                
                 FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
                End If
            
            Case "Dólares"
             Fecha = FrmSolicitudPagos.TxtFecha.Value
                      Fechas = Format(FrmSolicitudPagos.TxtFecha.Value, "yyyy/mm/dd")
'                      FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas & "', 102))"
                      FrmSolicitudPagos.DtaTasas.Refresh
             If Not FrmSolicitudPagos.DtaTasas.Recordset.EOF Then
                MontoTasa = FrmSolicitudPagos.DtaTasas.Recordset("MontoCordobas")
                If MontoTasa = 0 Then
                  MontoTasa = 1
                End If
            
               Select Case FrmSolicitudPagos.CmbMoneda.Text
                  Case "Córdobas"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                  Case "Dólares"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
                  Case "Libras"
                    MontoTasa = FrmSolicitudPagos.DtaTasas.Recordset("MontoLibras")
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)

                    
                 End Select
                Else
                  FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
               End If
            
            Case "Libras"
                      Fecha = FrmSolicitudPagos.TxtFecha.Value
                                            Fechas = Format(FrmSolicitudPagos.TxtFecha.Value, "yyyy/mm/dd")
'                      FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas & "', 102))"
                      FrmSolicitudPagos.DtaTasas.Refresh
                If Not FrmSolicitudPagos.DtaTasas.Recordset.EOF Then
                MontoTasa = FrmSolicitudPagos.DtaTasas.Recordset("MontoLibras")
               Select Case FrmSolicitudPagos.CmbMoneda.Text
                  Case "Córdobas"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Dólares"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Libras"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1

                    
                 End Select
                Else
                 FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
                End If
         
         End Select
       End If
         
         
         
         If FrmSolicitudPagos.CmbMoneda.Text = "Dólares" Then
            FrmSolicitudPagos.DtaIndice.Recordset("TipoMoneda") = "Dólares"
          Else
            FrmSolicitudPagos.DtaIndice.Recordset("TipoMoneda") = "Córdobas"
          End If
         
         FrmSolicitudPagos.DtaIndice.Recordset.Update
         Else
       Criterio = "CodCuentas='" & FrmSolicitudPagos.DBGTransacciones.Columns(0).Text & "'"
       FrmSolicitudPagos.DtaCuentas.Recordset.Find (Criterio)
       If Not FrmSolicitudPagos.DtaCuentas.Recordset.EOF Then
        TipoMoneda = FrmSolicitudPagos.DtaCuentas.Recordset("TipoMoneda")

         Select Case TipoMoneda
            Case "Córdobas"
                      Fecha = FrmSolicitudPagos.TxtFecha.Value
                       Fechas = Format(FrmSolicitudPagos.TxtFecha.Value, "yyyy/mm/dd")
'                      FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas & "', 102))"
                      FrmSolicitudPagos.DtaTasas.Refresh
                If Not FrmSolicitudPagos.DtaTasas.Recordset.EOF Then
                 MontoTasa = FrmSolicitudPagos.DtaTasas.Recordset("MontoCordobas")
                 If MontoTasa = 0 Then
                   MontoTasa = 1
                 End If
                 Select Case FrmSolicitudPagos.CmbMoneda.Text
                  Case "Córdobas"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
                  Case "Dólares"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Libras"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = MontoTasa
                   ' FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                    
                 End Select
                Else
                
                 FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
                End If
            
            Case "Dólares"
             Fecha = FrmSolicitudPagos.TxtFecha.Value
             FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
             FrmSolicitudPagos.DtaTasas.Refresh
             If Not FrmSolicitudPagos.DtaTasas.Recordset.EOF Then
                MontoTasa = FrmSolicitudPagos.DtaTasas.Recordset("MontoCordobas")
                If MontoTasa = 0 Then
                  MontoTasa = 1
                End If
            
               Select Case FrmSolicitudPagos.CmbMoneda.Text
                  Case "Córdobas"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                  Case "Dólares"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
                  Case "Libras"
                    MontoTasa = FrmSolicitudPagos.DtaTasas.Recordset("MontoLibras")
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)

                    
                 End Select
                Else
                  FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
               End If
            
            Case "Libras"
                      Fecha = FrmSolicitudPagos.TxtFecha.Value
                                            Fechas = Format(FrmSolicitudPagos.TxtFecha.Value, "yyyy/mm/dd")
'                      FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas & "', 102))"
                      FrmSolicitudPagos.DtaTasas.Refresh
                If Not FrmSolicitudPagos.DtaTasas.Recordset.EOF Then
                MontoTasa = FrmSolicitudPagos.DtaTasas.Recordset("MontoLibras")
               Select Case FrmSolicitudPagos.CmbMoneda.Text
                  Case "Córdobas"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Dólares"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Libras"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1

                    
                 End Select
                Else
                 FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
                End If
         
         End Select
       End If
           
         
         End If
        End If
 FrmSolicitudPagos.DtaCuentas.Refresh
  Criterio = "CodCuentas='" & FrmSolicitudPagos.DBCodigo.Text & "'"
  FrmSolicitudPagos.DtaCuentas.Recordset.Find (Criterio)
        
   TipoCuenta = FrmSolicitudPagos.DtaCuentas.Recordset("TipoCuenta")
   CodigoCuenta = FrmSolicitudPagos.DtaCuentas.Recordset("CodCuentas")
  If TipoCuenta = "Bancos" Then

   If Primero = True Then
     Primero = False
        Me.DtaConsulta.RecordSource = "SELECT NConsecutivoVoucher.CodCuenta, NConsecutivoVoucher.ConsecutivoVoucher, NConsecutivoVoucher.NPeriodo From NConsecutivoVoucher Where (((NConsecutivoVoucher.CodCuenta) = '" & CodigoCuenta & "') And ((NConsecutivoVoucher.NPeriodo) = " & NumeroPeriodo & "))"
        Me.DtaConsulta.Refresh
        If DtaConsulta.Recordset.EOF Then
           Me.DtaConsulta.Recordset.AddNew
             Me.DtaConsulta.Recordset("CodCuenta") = CodigoCuenta
             Me.DtaConsulta.Recordset("NPeriodo") = NumeroPeriodo
             Me.DtaConsulta.Recordset("ConsecutivoVoucher") = 1
           Me.DtaConsulta.Recordset.Update
           NumeroVoucher = 1
        Else
           'Me.'DtaConsulta.Recordset.Edit
            Me.DtaConsulta.Recordset("ConsecutivoVoucher") = Me.DtaConsulta.Recordset("ConsecutivoVoucher") + 1
           Me.DtaConsulta.Recordset.Update
         NumeroVoucher = Me.DtaConsulta.Recordset("ConsecutivoVoucher")
        End If
     Else
       ' FrmSolicitudPagos.DtaTransacciones.Recordset.MoveLast
     
     End If
        ConsecutivoVoucher = Month(FrmSolicitudPagos.TxtFecha.Value)
        If TipoCuenta = "Caja" Then
              numero = "CASH " & NumeroVoucher & "/" & ConsecutivoVoucher
        End If
        Select Case TipoMoneda
           Case "Córdobas"
            If TipoCuenta = "Bancos" Then
              numero = "BC " & NumeroVoucher & "/" & ConsecutivoVoucher
            End If
           Case "Dólares"
            If TipoCuenta = "Bancos" Then
              numero = "BD " & NumeroVoucher & "/" & ConsecutivoVoucher
            End If
        
         End Select
        
     End If
        ' Cadena = Mid(FrmSolicitudPagos.DBCodigo, 1, 1)
        ' Cadena = Cadena & "/" & NumeroTransaccion
        
   '///////////////////////////////////////////////////////////
   '//////CON ESTA CONSULTA BUSCO LA DESCRIPCION DE LA LINEA ANTERIOR//////////////////
   '/////////////////////////////////////////////////////////////////////////////////
   
            
            Sql = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta AS DescripcionCuentas, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, " & _
            "Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, " & _
            "Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, " & _
            "Transacciones.NumeroMovimiento , Periodos.Periodo " & _
            "FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo " & _
            "WHERE  (Transacciones.FechaTransaccion BETWEEN '" & Format(FrmSolicitudPagos.TxtFecha.Value, "yyyymmdd") & "' And '" & Format(FrmSolicitudPagos.TxtFecha.Value, "yyyymmdd") & "') AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ") " & _
            "ORDER BY Transacciones.NTransaccion"
              
            Me.DtaConsulta.RecordSource = Sql
            Me.DtaConsulta.Refresh
            If Not Me.DtaConsulta.Recordset.EOF Then
              Me.DtaConsulta.Recordset.MoveLast
              If Not IsNull(Me.DtaConsulta.Recordset("DescripcionMovimiento")) Then
                 DescripcionMovimiento = Me.DtaConsulta.Recordset("DescripcionMovimiento")
              End If
              If Not IsNull(Me.DtaConsulta.Recordset("Clave")) Then
                ClaveMovimiento = Me.DtaConsulta.Recordset("Clave")
              End If
            
            End If
        
        FrmSolicitudPagos.DBGTransacciones.Columns(3).Text = DescripcionMovimiento
        FrmSolicitudPagos.DBGTransacciones.Columns(2).Text = cadena
        If ClaveMovimiento = "" Then
         FrmSolicitudPagos.DBGTransacciones.Columns(6).Text = "Debito"
        Else
         FrmSolicitudPagos.DBGTransacciones.Columns(6).Text = ClaveMovimiento
        End If
        'FrmSolicitudPagos.DBGTransacciones.Columns(9).Locked = True
        FrmSolicitudPagos.DBGTransacciones.Columns(1).Text = rs("DescripcionCuentas")          'FrmSolicitudPagos.DtaCuentas.Recordset("DescripcionCuentas")
        FrmSolicitudPagos.DBGTransacciones.Columns(10).Text = Format(FrmSolicitudPagos.TxtFecha.Value, "dd/mm/yyyy")
        FrmSolicitudPagos.DBGTransacciones.Columns(11).Text = NumeroPeriodo
        FrmSolicitudPagos.DBGTransacciones.Columns(13).Text = FrmSolicitudPagos.TxtFuente.Text
        FrmSolicitudPagos.DBGTransacciones.Columns(14).Text = Format(FrmSolicitudPagos.TxtFecha.Value, "dd/mm/yyyy")
        FrmSolicitudPagos.DBGTransacciones.Columns(15).Text = NumeroTransaccion
         
'         For I = 2 To 5
'            If FrmSolicitudPagos.DBGTransacciones.Columns(I).Text = "" Then FrmSolicitudPagos.DBGTransacciones.Columns(I).Text = "-"
'        Next I
         
         'prueba de parche
'        FrmSolicitudPagos.DtaTransacciones.Refresh
'        inputFrmEgresos.DtaTransacciones.RecordSource
'         FrmSolicitudPagos.DBGTransacciones.Update
'
       Else
         MsgBox "La cuenta digitada no es correcta", vbCritical, "Sistema Contable"
         Exit Sub
       End If

End Sub


Private Sub PegarSolicitudCheque()

    FrmSolicitudPagos.DtaCuentas.Refresh
    FrmSolicitudPagos.DBGTransacciones.Columns(0).Text = rs("CodCuentas")
     Criterio = "CodCuentas='" & FrmSolicitudPagos.DBGTransacciones.Columns(0).Text & "'"
      FrmSolicitudPagos.DtaCuentas.Recordset.Find (Criterio)
       If Not FrmSolicitudPagos.DtaCuentas.Recordset.EOF Then
         mes = Month(FrmSolicitudPagos.TxtFecha.Value)
         Año = Year(FrmSolicitudPagos.TxtFecha.Value)
         FechaIni = CDate("1/" & Month(FrmSolicitudPagos.TxtFecha.Value) & "/" & Year(FrmSolicitudPagos.TxtFecha.Value))
         FechaFin = DateSerial(Año, mes + 1, 1 - 1)
         NumFecha1 = CDate(FechaIni)
         NumFecha2 = CDate(FechaFin)
 
        FrmSolicitudPagos.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
        FrmSolicitudPagos.DtaConsulta.Refresh
         If Not FrmSolicitudPagos.DtaConsulta.Recordset.EOF Then
          FrmSolicitudPagos.TxtPeriodo.Text = FrmSolicitudPagos.DtaConsulta.Recordset("Periodo")
            NumeroPeriodo = FrmSolicitudPagos.DtaConsulta.Recordset("NPeriodo")
            If Val(FrmSolicitudPagos.TxtNTransacciones.Text) = 0 Then
                NumeroTransaccion = FrmSolicitudPagos.NumeroSolicitud
            Else
                NumeroTransaccion = FrmSolicitudPagos.NumeroSolicitud
            End If
            EstadoPeriodo = FrmSolicitudPagos.DtaConsulta.Recordset("EstadoPeriodo")
      
        '////////////Edito los datos del Periodo///////////
                 If Val(FrmSolicitudPagos.TxtNTransacciones.Text) = 0 Then
                  
                  
                        Me.AdoBusca.RecordSource = "SELECT NConsecutivos.CodCuentas, NConsecutivos.ConsecutivoSolicitudCheque  From NConsecutivos Where (((NConsecutivos.CodCuentas) = '" & FrmSolicitudPagos.DBCodigo.Text & "'))"
                        Me.AdoBusca.Refresh
                        If Me.AdoBusca.Recordset.EOF Then
                           Me.AdoBusca.Recordset.AddNew
                             Me.AdoBusca.Recordset("CodCuentas") = FrmSolicitudPagos.DBCodigo.Text
                             Me.AdoBusca.Recordset("ConsecutivoSolicitudCheque") = NumeroTransaccion
                           Me.AdoBusca.Recordset.Update
                
                          
                        Else
                        
                           Me.AdoBusca.Recordset("ConsecutivoSolicitudCheque") = NumeroTransaccion
                           Me.AdoBusca.Recordset.Update
                
                       End If
                       
                      NumeroTransaccion = Format(FrmSolicitudPagos.NumeroSolicitud, "0000#")
                      FrmSolicitudPagos.TxtNTransacciones.Text = Format(NumeroSolicitud, "0000#")

                  '////////Edito los Datos de los indices de Transacciones//////
                 
                 FrmSolicitudPagos.DtaIndice.Recordset.AddNew
                 FrmSolicitudPagos.DtaIndice.Recordset("FechaTransaccion") = Format(FrmSolicitudPagos.TxtFecha.Value, "dd/mm/yyyy")
                 FrmSolicitudPagos.DtaIndice.Recordset("NumeroMovimiento") = NumeroTransaccion
                 FrmSolicitudPagos.DtaIndice.Recordset("Fuente") = FrmSolicitudPagos.TxtFuente.Text
                 FrmSolicitudPagos.DtaIndice.Recordset("NPeriodo") = NumeroPeriodo
                 
              Criterio = "CodCuentas='" & FrmSolicitudPagos.DBGTransacciones.Columns(0).Text & "'"
               FrmSolicitudPagos.DtaCuentas.Recordset.Find (Criterio)
               If Not FrmSolicitudPagos.DtaCuentas.Recordset.EOF Then
                TipoMoneda = FrmSolicitudPagos.DtaCuentas.Recordset("TipoMoneda")
        
                 Select Case TipoMoneda
                    Case "Córdobas"
                    
                              Fecha = FrmSolicitudPagos.TxtFecha.Value
                              Fechas = Format(FrmSolicitudPagos.TxtFecha.Value, "yyyy/mm/dd")
        '                      FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                              FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE (FechaTasas = '" & Format(FrmSolicitudPagos.TxtFecha.Value, "yyyymmdd") & "')"
                              FrmSolicitudPagos.DtaTasas.Refresh
                        If Not FrmSolicitudPagos.DtaTasas.Recordset.EOF Then
                         MontoTasa = FrmSolicitudPagos.DtaTasas.Recordset("MontoCordobas")
                         If MontoTasa = 0 Then
                           MontoTasa = 1
                         End If
                         Select Case FrmSolicitudPagos.CmbMoneda.Text
                          Case "Córdobas"
                            FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
                          Case "Dólares"
                            FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = MontoTasa
                          Case "Libras"
                            FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = MontoTasa
                           ' FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                            
                         End Select
                        Else
                        
                         FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
                        End If
                    
                    Case "Dólares"
                     Fecha = FrmSolicitudPagos.TxtFecha.Value
                              Fechas = Format(FrmSolicitudPagos.TxtFecha.Value, "yyyy/mm/dd")
        '                      FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                              FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas & "', 102))"
                              FrmSolicitudPagos.DtaTasas.Refresh
                     If Not FrmSolicitudPagos.DtaTasas.Recordset.EOF Then
                        MontoTasa = FrmSolicitudPagos.DtaTasas.Recordset("MontoCordobas")
                        If MontoTasa = 0 Then
                          MontoTasa = 1
                        End If
                    
                       Select Case FrmSolicitudPagos.CmbMoneda.Text
                          Case "Córdobas"
                            FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                          Case "Dólares"
                            FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
                          Case "Libras"
                            MontoTasa = FrmSolicitudPagos.DtaTasas.Recordset("MontoLibras")
                            FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
        
                            
                         End Select
                        Else
                          FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
                       End If
                    
                    Case "Libras"
                              Fecha = FrmSolicitudPagos.TxtFecha.Value
                                                    Fechas = Format(FrmSolicitudPagos.TxtFecha.Value, "yyyy/mm/dd")
        '                      FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                              FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas & "', 102))"
                              FrmSolicitudPagos.DtaTasas.Refresh
                        If Not FrmSolicitudPagos.DtaTasas.Recordset.EOF Then
                        MontoTasa = FrmSolicitudPagos.DtaTasas.Recordset("MontoLibras")
                       Select Case FrmSolicitudPagos.CmbMoneda.Text
                          Case "Córdobas"
                            FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = MontoTasa
                          Case "Dólares"
                            FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = MontoTasa
                          Case "Libras"
                            FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
        
                            
                         End Select
                        Else
                         FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
                        End If
                 
                 End Select
               End If
         
         
         
            If FrmSolicitudPagos.CmbMoneda.Text = "Dólares" Then
              FrmSolicitudPagos.DtaIndice.Recordset("TipoMoneda") = "Dólares"
            Else
              FrmSolicitudPagos.DtaIndice.Recordset("TipoMoneda") = "Córdobas"
            End If
            
          If FrmSolicitudPagos.ChkCheque.Value = 1 Then
            FrmSolicitudPagos.DtaIndice.Recordset("ImprimeCheque") = 1
          Else
            FrmSolicitudPagos.DtaIndice.Recordset("ImprimeCheque") = 0
          End If
          
          If FrmSolicitudPagos.ChkCheque.Value = 1 Then
            FrmSolicitudPagos.DtaIndice.Recordset("Retencion1") = 1
          Else
            FrmSolicitudPagos.DtaIndice.Recordset("Retencion1") = 0
          End If
          
          If FrmSolicitudPagos.Chk2.Value = 1 Then
            FrmSolicitudPagos.DtaIndice.Recordset("Retencion2") = 1
          Else
            FrmSolicitudPagos.DtaIndice.Recordset("Retencion2") = 0
          End If
          
          If FrmSolicitudPagos.Chk3.Value = 1 Then
            FrmSolicitudPagos.DtaIndice.Recordset("Retencion3") = 1
          Else
            FrmSolicitudPagos.DtaIndice.Recordset("Retencion3") = 0
          End If
          
          If FrmSolicitudPagos.Chk7.Value = 1 Then
            FrmSolicitudPagos.DtaIndice.Recordset("Retencion4") = 1
          Else
            FrmSolicitudPagos.DtaIndice.Recordset("Retencion4") = 0
          End If
          
          If FrmSolicitudPagos.Chk10.Value = 1 Then
            FrmSolicitudPagos.DtaIndice.Recordset("Retencion5") = 1
          Else
            FrmSolicitudPagos.DtaIndice.Recordset("Retencion5") = 0
          End If
          
          If FrmSolicitudPagos.Chk15.Value = 1 Then
            FrmSolicitudPagos.DtaIndice.Recordset("Iva") = 1
          Else
            FrmSolicitudPagos.DtaIndice.Recordset("Iva") = 0
          End If
          
          FrmSolicitudPagos.DtaIndice.Recordset("Concepto") = FrmSolicitudPagos.TxtMemo.Text
          FrmSolicitudPagos.DtaIndice.Recordset("SubTotal") = FrmSolicitudPagos.TxtSubTotal.Text
          FrmSolicitudPagos.DtaIndice.Recordset("MontoIva") = FrmSolicitudPagos.TxtIVa.Text
          FrmSolicitudPagos.DtaIndice.Recordset("MontoRetenciones") = FrmSolicitudPagos.TxtRetenciones.Text
          FrmSolicitudPagos.DtaIndice.Recordset("MontoSolicitud") = FrmSolicitudPagos.TxtMonto.Text
          FrmSolicitudPagos.DtaIndice.Recordset("Beneficiario") = FrmSolicitudPagos.TxtNombre.Text
          FrmSolicitudPagos.DtaIndice.Recordset("CuentaBanco") = FrmSolicitudPagos.DBCodigo.Text
         
         FrmSolicitudPagos.DtaIndice.Recordset.Update
     Else
       Criterio = "CodCuentas='" & FrmSolicitudPagos.DBGTransacciones.Columns(0).Text & "'"
       FrmSolicitudPagos.DtaCuentas.Recordset.Find (Criterio)
       If Not FrmSolicitudPagos.DtaCuentas.Recordset.EOF Then
        TipoMoneda = FrmSolicitudPagos.DtaCuentas.Recordset("TipoMoneda")

         Select Case TipoMoneda
            Case "Córdobas"
                      Fecha = FrmSolicitudPagos.TxtFecha.Value
                       Fechas = Format(FrmSolicitudPagos.TxtFecha.Value, "yyyy/mm/dd")
'                      FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas & "', 102))"
                      FrmSolicitudPagos.DtaTasas.Refresh
                If Not FrmSolicitudPagos.DtaTasas.Recordset.EOF Then
                 MontoTasa = FrmSolicitudPagos.DtaTasas.Recordset("MontoCordobas")
                 If MontoTasa = 0 Then
                   MontoTasa = 1
                 End If
                 Select Case FrmSolicitudPagos.CmbMoneda.Text
                  Case "Córdobas"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
                  Case "Dólares"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Libras"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = MontoTasa
                   ' FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                    
                 End Select
                Else
                
                 FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
                End If
            
            Case "Dólares"
             Fecha = FrmSolicitudPagos.TxtFecha.Value
             FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
             FrmSolicitudPagos.DtaTasas.Refresh
             If Not FrmSolicitudPagos.DtaTasas.Recordset.EOF Then
                MontoTasa = FrmSolicitudPagos.DtaTasas.Recordset("MontoCordobas")
                If MontoTasa = 0 Then
                  MontoTasa = 1
                End If
            
               Select Case FrmSolicitudPagos.CmbMoneda.Text
                  Case "Córdobas"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)
                  Case "Dólares"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
                  Case "Libras"
                    MontoTasa = FrmSolicitudPagos.DtaTasas.Recordset("MontoLibras")
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = (1 / MontoTasa)

                    
                 End Select
                Else
                  FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
               End If
            
            Case "Libras"
                      Fecha = FrmSolicitudPagos.TxtFecha.Value
                                            Fechas = Format(FrmSolicitudPagos.TxtFecha.Value, "yyyy/mm/dd")
'                      FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) =" & Fecha & "))"
                      FrmSolicitudPagos.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas & "', 102))"
                      FrmSolicitudPagos.DtaTasas.Refresh
                If Not FrmSolicitudPagos.DtaTasas.Recordset.EOF Then
                MontoTasa = FrmSolicitudPagos.DtaTasas.Recordset("MontoLibras")
               Select Case FrmSolicitudPagos.CmbMoneda.Text
                  Case "Córdobas"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Dólares"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = MontoTasa
                  Case "Libras"
                    FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1

                    
                 End Select
                Else
                 FrmSolicitudPagos.DBGTransacciones.Columns(7).Text = 1
                End If
         
         End Select
       End If
           
         
         End If
        End If
 FrmSolicitudPagos.DtaCuentas.Refresh
  Criterio = "CodCuentas='" & FrmSolicitudPagos.DBCodigo.Text & "'"
  FrmSolicitudPagos.DtaCuentas.Recordset.Find (Criterio)
        
   TipoCuenta = FrmSolicitudPagos.DtaCuentas.Recordset("TipoCuenta")
   CodigoCuenta = FrmSolicitudPagos.DtaCuentas.Recordset("CodCuentas")
  If TipoCuenta = "Bancos" Then

   If Primero = True Then
     Primero = False
        Me.DtaConsulta.RecordSource = "SELECT NConsecutivoVoucher.CodCuenta, NConsecutivoVoucher.ConsecutivoVoucher, NConsecutivoVoucher.NPeriodo From NConsecutivoVoucher Where (((NConsecutivoVoucher.CodCuenta) = '" & CodigoCuenta & "') And ((NConsecutivoVoucher.NPeriodo) = " & NumeroPeriodo & "))"
        Me.DtaConsulta.Refresh
        If DtaConsulta.Recordset.EOF Then
           Me.DtaConsulta.Recordset.AddNew
             Me.DtaConsulta.Recordset("CodCuenta") = CodigoCuenta
             Me.DtaConsulta.Recordset("NPeriodo") = NumeroPeriodo
             Me.DtaConsulta.Recordset("ConsecutivoVoucher") = 1
           Me.DtaConsulta.Recordset.Update
           NumeroVoucher = 1
        Else
           'Me.'DtaConsulta.Recordset.Edit
            Me.DtaConsulta.Recordset("ConsecutivoVoucher") = Me.DtaConsulta.Recordset("ConsecutivoVoucher") + 1
           Me.DtaConsulta.Recordset.Update
         NumeroVoucher = Me.DtaConsulta.Recordset("ConsecutivoVoucher")
        End If
     Else
       ' FrmSolicitudPagos.DtaTransacciones.Recordset.MoveLast
     
     End If
        ConsecutivoVoucher = Month(FrmSolicitudPagos.TxtFecha.Value)
        If TipoCuenta = "Caja" Then
              numero = "CASH " & NumeroVoucher & "/" & ConsecutivoVoucher
        End If
        Select Case TipoMoneda
           Case "Córdobas"
            If TipoCuenta = "Bancos" Then
              numero = "BC " & NumeroVoucher & "/" & ConsecutivoVoucher
            End If
           Case "Dólares"
            If TipoCuenta = "Bancos" Then
              numero = "BD " & NumeroVoucher & "/" & ConsecutivoVoucher
            End If
        
         End Select
        
     End If
        ' Cadena = Mid(FrmSolicitudPagos.DBCodigo, 1, 1)
        ' Cadena = Cadena & "/" & NumeroTransaccion
        
   '///////////////////////////////////////////////////////////
   '//////CON ESTA CONSULTA BUSCO LA DESCRIPCION DE LA LINEA ANTERIOR//////////////////
   '/////////////////////////////////////////////////////////////////////////////////
   
            
            Sql = "SELECT TransaccionesSolicitudPago.CodCuentas, TransaccionesSolicitudPago.NombreCuenta AS DescripcionCuentas, TransaccionesSolicitudPago.VoucherNo, TransaccionesSolicitudPago.DescripcionMovimiento, " & _
            "TransaccionesSolicitudPago.FacturaNo, TransaccionesSolicitudPago.ChequeNo, TransaccionesSolicitudPago.Clave, TransaccionesSolicitudPago.TCambio, TransaccionesSolicitudPago.Debito, TransaccionesSolicitudPago.Credito, " & _
            "TransaccionesSolicitudPago.FechaTransaccion, TransaccionesSolicitudPago.NPeriodo, TransaccionesSolicitudPago.NTransaccion, TransaccionesSolicitudPago.Fuente, TransaccionesSolicitudPago.FechaTasas, " & _
            "TransaccionesSolicitudPago.NumeroMovimiento , Periodos.Periodo " & _
            "FROM Periodos INNER JOIN TransaccionesSolicitudPago ON Periodos.NPeriodo = TransaccionesSolicitudPago.NPeriodo " & _
            "WHERE  (TransaccionesSolicitudPago.FechaTransaccion BETWEEN '" & Format(FrmSolicitudPagos.TxtFecha.Value, "yyyymmdd") & "' And '" & Format(FrmSolicitudPagos.TxtFecha.Value, "yyyymmdd") & "') AND (TransaccionesSolicitudPago.NumeroMovimiento = " & NumeroTransaccion & ") " & _
            "ORDER BY TransaccionesSolicitudPago.NTransaccion"
              
            Me.DtaConsulta.RecordSource = Sql
            Me.DtaConsulta.Refresh
            If Not Me.DtaConsulta.Recordset.EOF Then
              Me.DtaConsulta.Recordset.MoveLast
              If Not IsNull(Me.DtaConsulta.Recordset("DescripcionMovimiento")) Then
                 DescripcionMovimiento = Me.DtaConsulta.Recordset("DescripcionMovimiento")
              End If
              If Not IsNull(Me.DtaConsulta.Recordset("Clave")) Then
                ClaveMovimiento = Me.DtaConsulta.Recordset("Clave")
              End If
            
            End If
        
        FrmSolicitudPagos.DBGTransacciones.Columns(3).Text = DescripcionMovimiento
        FrmSolicitudPagos.DBGTransacciones.Columns(2).Text = cadena
        If ClaveMovimiento = "" Then
         FrmSolicitudPagos.DBGTransacciones.Columns(6).Text = "Debito"
        Else
         FrmSolicitudPagos.DBGTransacciones.Columns(6).Text = ClaveMovimiento
        End If
        'FrmSolicitudPagos.DBGTransacciones.Columns(9).Locked = True
        FrmSolicitudPagos.DBGTransacciones.Columns(1).Text = rs("DescripcionCuentas")          'FrmSolicitudPagos.DtaCuentas.Recordset("DescripcionCuentas")
        FrmSolicitudPagos.DBGTransacciones.Columns(10).Text = Format(FrmSolicitudPagos.TxtFecha.Value, "dd/mm/yyyy")
        FrmSolicitudPagos.DBGTransacciones.Columns(11).Text = NumeroPeriodo
        FrmSolicitudPagos.DBGTransacciones.Columns(13).Text = FrmSolicitudPagos.TxtFuente.Text
        FrmSolicitudPagos.DBGTransacciones.Columns(14).Text = Format(FrmSolicitudPagos.TxtFecha.Value, "dd/mm/yyyy")
        FrmSolicitudPagos.DBGTransacciones.Columns(15).Text = NumeroTransaccion
         
'         For I = 2 To 5
'            If FrmSolicitudPagos.DBGTransacciones.Columns(I).Text = "" Then FrmSolicitudPagos.DBGTransacciones.Columns(I).Text = "-"
'        Next I
         
         'prueba de parche
'        FrmSolicitudPagos.DtaTransacciones.Refresh
'        inputFrmEgresos.DtaTransacciones.RecordSource
'         FrmSolicitudPagos.DBGTransacciones.Update
'
       Else
         MsgBox "La cuenta digitada no es correcta", vbCritical, "Sistema Contable"
         Exit Sub
       End If

End Sub





