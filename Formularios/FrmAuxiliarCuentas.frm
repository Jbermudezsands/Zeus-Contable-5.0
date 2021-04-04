VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FrmAuxiliarCuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auxiliar de Cuentas"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12600
   Icon            =   "FrmAuxiliarCuentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   12600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdConciliacion 
      Caption         =   "Conciliacion"
      Height          =   375
      Left            =   5400
      TabIndex        =   15
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   11400
      TabIndex        =   16
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   4080
      TabIndex        =   14
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton CmdSiguiente 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "Anterior"
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton CmdMovimientos 
      Caption         =   "Ver Mov."
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   7680
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblSaldoCuenta 
      Height          =   255
      Left            =   1800
      OleObjectBlob   =   "FrmAuxiliarCuentas.frx":030A
      TabIndex        =   10
      Top             =   1560
      Width           =   3975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmAuxiliarCuentas.frx":0368
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DBCliente 
      Bindings        =   "FrmAuxiliarCuentas.frx":03EC
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CodCuentas"
      Text            =   ""
   End
   Begin TrueOleDBGrid80.TDBGrid DBGCuentas 
      Bindings        =   "FrmAuxiliarCuentas.frx":0405
      Height          =   5535
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   9763
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
      Splits(0).Caption=   "Tarjeta Auxiliar de Cuentas"
      Splits(0).DividerColor=   14215660
      Splits(0).FilterBar=   -1  'True
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
      AllowUpdate     =   0   'False
      Appearance      =   3
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      PictureCurrentRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      PictureCurrentRow(0)=   "bHQAAOYBAABCTeYBAAAAAAAANgAAACgAAAAPAAAACQAAAAEAGAAAAAAAsAEAAAAAAAAAAAAAAAAA"
      PictureCurrentRow(1)=   "AAAAAAD///////////////////////////////////////////////////////////8AAAD/////"
      PictureCurrentRow(2)=   "//////////////////////////////////////////////////////8AAAD///////8AhgAAhgAA"
      PictureCurrentRow(3)=   "hgAAhgAAhgAAhgAAhgAAhgAAhgAAhgAAhgD///////8AAAD///////8AhgD///+EhoSEhoSEhoSE"
      PictureCurrentRow(4)=   "hoSEhoSEhoSEhoSEhoQAhgD///////8AAAD///////8AhgD////Gx8bGx8bGx8bGx8bGx8bGx8bG"
      PictureCurrentRow(5)=   "x8aEhoQAhgD///////8AAAD///////8AhgD///////////////////////////////////8AhgD/"
      PictureCurrentRow(6)=   "//////8AAAD///////8AhgAAhgAAhgAAhgAAhgAAhgAAhgAAhgAAhgAAhgAAhgD///////8AAAD/"
      PictureCurrentRow(7)=   "//////////////////////////////////////////////////////////8AAAD/////////////"
      PictureCurrentRow(8)=   "//////////////////////////////////////////////8AAAA="
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
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HCE9D9D&,.bold=-1"
      _StyleDefs(20)  =   ":id=22,.fontsize=1200,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(21)  =   ":id=22,.fontname=Script MT Bold"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(23)  =   ":id=14,.fgcolor=&H0&"
      _StyleDefs(24)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(25)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(26)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(27)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(28)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(29)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(30)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(31)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(32)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(41)  =   "Named:id=33:Normal"
      _StyleDefs(42)  =   ":id=33,.parent=0"
      _StyleDefs(43)  =   "Named:id=34:Heading"
      _StyleDefs(44)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(45)  =   ":id=34,.wraptext=-1"
      _StyleDefs(46)  =   "Named:id=35:Footing"
      _StyleDefs(47)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   "Named:id=36:Selected"
      _StyleDefs(49)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(50)  =   "Named:id=37:Caption"
      _StyleDefs(51)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(52)  =   "Named:id=38:HighlightRow"
      _StyleDefs(53)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(54)  =   "Named:id=39:EvenRow"
      _StyleDefs(55)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(56)  =   "Named:id=40:OddRow"
      _StyleDefs(57)  =   ":id=40,.parent=33"
      _StyleDefs(58)  =   "Named:id=41:RecordSelector"
      _StyleDefs(59)  =   ":id=41,.parent=34"
      _StyleDefs(60)  =   "Named:id=42:FilterBar"
      _StyleDefs(61)  =   ":id=42,.parent=33"
   End
   Begin MSAdodcLib.Adodc DtaSaldos 
      Height          =   330
      Left            =   360
      Top             =   9960
      Width           =   3375
      _ExtentX        =   5953
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
   Begin MSAdodcLib.Adodc DtaConsulta 
      Height          =   375
      Left            =   360
      Top             =   9360
      Width           =   3375
      _ExtentX        =   5953
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
   Begin MSAdodcLib.Adodc DtaCuentas 
      Height          =   375
      Left            =   360
      Top             =   9600
      Width           =   3495
      _ExtentX        =   6165
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
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales de la Cuenta"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12135
      Begin VB.ComboBox CmbTipoMoneda 
         Height          =   315
         ItemData        =   "FrmAuxiliarCuentas.frx":041D
         Left            =   6840
         List            =   "FrmAuxiliarCuentas.frx":0427
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox TxtDescripcion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   2
         Top             =   720
         Width           =   10455
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
         Left            =   4320
         Picture         =   "FrmAuxiliarCuentas.frx":043E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   5520
         OleObjectBlob   =   "FrmAuxiliarCuentas.frx":058C
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmAuxiliarCuentas.frx":0600
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmAuxiliarCuentas.frx":066A
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.CommandButton CmdAuditoria 
      Caption         =   "Auditoria"
      Height          =   375
      Left            =   5400
      TabIndex        =   17
      Top             =   7680
      Width           =   1095
   End
End
Attribute VB_Name = "FrmAuxiliarCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnx As New ADODB.Connection
Private rs As New ADODB.Recordset, rsConexion As New ADODB.Recordset
Private SQL As String
Private modal As Boolean
Private getVal As Boolean
Private Id As Integer
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set cnx = Nothing
Set rs = Nothing
Set rs2 = Nothing
Set rs3 = Nothing
End Sub


Private Sub CmbTipoMoneda_Click()
On Error GoTo TipoErrs
Dim Debito As Double, Credito As Double
Total1 = 0
Me.DtaCuentas.Refresh
Criterio = "CodCuentas='" & Me.DBCliente.Text & "'"
Me.DtaCuentas.Recordset.Find (Criterio)
If DtaCuentas.Recordset.EOF Then
 
  sqlconsulta = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento,VoucherNo, FacturaNo, ChequeNo, Transacciones.DescripcionMovimiento, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Debito*TCambio AS MDebito, TCambio*Credito AS MCredito From Transacciones Where (((Transacciones.CodCuentas) =  '0')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
'  Me.DtaSaldos.Refresh
        rs.Close
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
       
       Me.DBGCuentas.DataSource = rs
       
  Me.DBGCuentas.Columns(1).Caption = "Fecha"
  Me.DBGCuentas.Columns(1).Width = 1000
  Me.DBGCuentas.Columns(2).Caption = "No.Trans"
  Me.DBGCuentas.Columns(2).Width = 1000
  Me.DBGCuentas.Columns(3).Width = 1000
  Me.DBGCuentas.Columns(4).Width = 1000
  Me.DBGCuentas.Columns(5).Width = 1000
  Me.DBGCuentas.Columns(5).Caption = "Cheq/Rec"
  Me.DBGCuentas.Columns(6).Width = 3000
  Me.DBGCuentas.Columns(7).Width = 1000
  Me.DBGCuentas.Columns(7).NumberFormat = "##,##0.000000"
  Me.DBGCuentas.Columns(8).Visible = False
  Me.DBGCuentas.Columns(9).Visible = False
  Me.DBGCuentas.Columns(10).Width = 1200
  Me.DBGCuentas.Columns(10).Caption = "Debito"
  Me.DBGCuentas.Columns(11).Width = 1200
  Me.DBGCuentas.Columns(11).Caption = "Credito"
  Me.DBGCuentas.Columns(10).NumberFormat = "##,##0.00"
  Me.DBGCuentas.Columns(11).NumberFormat = "##,##0.00"
  Me.LblSaldoCuenta.Caption = "0.00"
  Me.DBGCuentas.Columns(0).Visible = False
  Me.CmbTipoMoneda.Enabled = True
  
Else
If QUIEN = "Nuevo" Then
'      Me.TxtDescripcion.Text = Me.DtaCuentas.Recordset("DescripcionCuentas")
'      Me.CmbTipoMoneda.Text = Me.DtaCuentas.Recordset("TipoMoneda")
'      TipoCuenta = DtaCuentas.Recordset("TipoCuenta")
'      If Not IsNull(DtaCuentas.Recordset("TipoMoneda")) Then
'       Me.CmbTipoMoneda = DtaCuentas.Recordset("TipoMoneda")
'      End If
      QUIEN = "Viejo"
      Exit Sub
     ' Me.LblSaldoCuenta.Caption = Format(DtaCuentas.Recordset.SaldoActual, "##,##0.00")
 Else
     TipoCuenta = DtaCuentas.Recordset("TipoCuenta")
 End If

  '//////////Muestro los Saldos de las cuentas/////////////////////
  
 If Me.CmbTipoMoneda.Text = "Córdobas" Then
    sqlconsulta = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento, VoucherNo, FacturaNo, ChequeNo,Transacciones.DescripcionMovimiento, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Debito*TCambio AS MDebito, TCambio*Credito AS MCredito From Transacciones Where (((Transacciones.CodCuentas) =  '" & Me.DBCliente.Text & "')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
'    Me.DtaSaldos.Refresh
        rs.Close
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        
       Me.DBGCuentas.DataSource = rs
 Else
    sqlconsulta = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento, Transacciones.VoucherNo,Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.DescripcionMovimiento, Transacciones.TCambio, Transacciones.Debito,Transacciones.Credito, ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 3) AS MDebito, ROUND(Transacciones.Credito * Transacciones.TCambio / Tasas.MontoCordobas, 3) As MCredito FROM Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas WHERE (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "') ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
'    Me.DtaSaldos.Refresh
        rs.Close
        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        
       Me.DBGCuentas.DataSource = rs
 End If
 
 If Me.CmbTipoMoneda.Text = "Córdobas" Then
    Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento, Transacciones.DescripcionMovimiento, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, Debito*TCambio AS MDebito, TCambio*Credito AS MCredito From Transacciones Where (((Transacciones.CodCuentas) =  '" & Me.DBCliente.Text & "')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
    Me.DtaConsulta.Refresh
 Else
    Me.DtaConsulta.RecordSource = "SELECT  Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento, Transacciones.DescripcionMovimiento,Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito,ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 3) AS MDebito,ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas, 3) As MCredito FROM Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas WHERE (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "') ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
    Me.DtaConsulta.Refresh
 End If
  
  Do While Not Me.DtaConsulta.Recordset.EOF
'   Me.CmbTipoMoneda.Enabled = False
   If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
    If Not IsNull(Me.DtaConsulta.Recordset("MDebito")) Then
     Debito = Format(Me.DtaConsulta.Recordset("MDebito"), "##,##0.00")
    End If
    If Not IsNull(Me.DtaConsulta.Recordset("MCredito")) Then
     Credito = Format(Me.DtaConsulta.Recordset("MCredito"), "##,##0.00")
    End If
    Total1 = Debito - Credito + Total1
    Debito = 0
    Credito = 0
   Else
    If Not IsNull(Me.DtaConsulta.Recordset("MDebito")) Then
     Debito = Format(Me.DtaConsulta.Recordset("MDebito"), "##,##0.00")
    End If
    If Not IsNull(Me.DtaConsulta.Recordset("MCredito")) Then
     Credito = Format(Me.DtaConsulta.Recordset("MCredito"), "##,##0.00")
    End If
    Total1 = Credito - Debito + Total1
    Debito = 0
    Credito = 0
   End If
   
   Me.DtaConsulta.Recordset.MoveNext
  Loop
  
  If TipoCuenta = "Bancos" Then
   Me.CmdAuditoria.Visible = False
   Me.CmdConciliacion.Visible = True
  Else
   Me.CmdAuditoria.Visible = True
   Me.CmdConciliacion.Visible = False
  End If
  Me.DBGCuentas.Columns(1).Caption = "Fecha"
  Me.DBGCuentas.Columns(1).Width = 1000
  Me.DBGCuentas.Columns(2).Caption = "No.Trans"
  Me.DBGCuentas.Columns(2).Width = 1000
  Me.DBGCuentas.Columns(3).Width = 1000
  Me.DBGCuentas.Columns(4).Width = 1000
  Me.DBGCuentas.Columns(5).Width = 1000
  Me.DBGCuentas.Columns(5).Caption = "Cheq/Rec"
  Me.DBGCuentas.Columns(6).Width = 3000
  Me.DBGCuentas.Columns(7).Width = 1000
  Me.DBGCuentas.Columns(7).NumberFormat = "##,##0.000000"
  Me.DBGCuentas.Columns(8).Visible = False
  Me.DBGCuentas.Columns(9).Visible = False
  Me.DBGCuentas.Columns(10).Width = 1200
  Me.DBGCuentas.Columns(10).Caption = "Debito"
  Me.DBGCuentas.Columns(11).Width = 1200
  Me.DBGCuentas.Columns(11).Caption = "Credito"
  Me.DBGCuentas.Columns(10).NumberFormat = "##,##0.00"
  Me.DBGCuentas.Columns(11).NumberFormat = "##,##0.00"
  
  Me.LblSaldoCuenta.Caption = Format(Total1, "##,##0.00")
  Me.DBGCuentas.Columns(0).Visible = False
 
  
 
End If
Exit Sub
TipoErrs:
ControlErrores
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

Private Sub CmdAuditoria_Click()
FrmConciliacion.Caption = "Auditoria de Cuentas"
FrmConciliacion.Show 1
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
   Me.DtaSaldos.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento, Transacciones.DescripcionMovimiento, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.Debito*Transacciones.TCambio AS MDebito, Transacciones.TCambio*Transacciones.Credito AS MCredito From Transacciones Where (((Transacciones.CodCuentas) =  '" & Me.DBCliente.Text & "')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
   Me.DtaSaldos.Refresh
  Me.DBGCuentas.Columns(2).Caption = "No.Trans"
  Me.DBGCuentas.Columns(2).Width = 1000
  Me.DBGCuentas.Columns(3).Width = 3000
   Me.DBGCuentas.Columns(0).Visible = False
  Me.DBGCuentas.Columns(5).Visible = False
  Me.DBGCuentas.Columns(6).Visible = False
  Me.DBGCuentas.Columns(7).Caption = "Debito"
  Me.DBGCuentas.Columns(8).Caption = "Credito"
  Me.DBGCuentas.Columns(7).NumberFormat = "##,##0.00"
  Me.DBGCuentas.Columns(8).NumberFormat = "##,##0.00"
  Me.DBGCuentas.Columns(4).NumberFormat = "##,##0.000000"
  
  Me.LblSaldoCuenta.Caption = Format(Total1, "##,##0.00")
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
  Me.DBGCuentas.Columns(2).Caption = "No.Trans"
  Me.DBGCuentas.Columns(2).Width = 1000
  Me.DBGCuentas.Columns(3).Width = 3000
  Me.DBGCuentas.Columns(5).Visible = False
  Me.DBGCuentas.Columns(6).Visible = False
  Me.DBGCuentas.Columns(7).Caption = "Debito"
  Me.DBGCuentas.Columns(8).Caption = "Credito"
  Me.DBGCuentas.Columns(7).NumberFormat = "##,##0.00"
  Me.DBGCuentas.Columns(8).NumberFormat = "##,##0.00"
  Me.DBGCuentas.Columns(4).NumberFormat = "##,##0.000000"
  Me.LblSaldoCuenta.Caption = Format(Total1, "##,##0.00")
 Me.DBGCuentas.Columns(0).Visible = False


 Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub CmdBuscaCuenta_Click()

     
QueProducto = "Auxiliar"
FrmConsulta.Show 1

End Sub

Private Sub CmdConciliacion_Click()


FrmConciliacion.Show 1



End Sub

Private Sub CmdMovimientos_Click()
Dim Fechas1 As String, Fechas2 As String
If Me.DBGCuentas.Columns(1).Text = "" Then
 Exit Sub
End If
FrmAuxiliarMovimientos.TxtFecha.Value = Me.DBGCuentas.Columns(1).Text
NumeroTransaccion = Me.DBGCuentas.Columns(2).Text
  If Not Me.DBGCuentas.Columns(3).Text = "" Then
    Descripcion = Me.DBGCuentas.Columns(3).Text
   If Descripcion = "**********CANCELADO*************" Then
     MsgBox " Este movimiento esta Cancelado", vbCritical, "Sistema Contable"
     Exit Sub
   End If
  End If
   
    FrmAuxiliarMovimientos.Enabled = True
 FrmAuxiliarMovimientos.CmbMoneda.Enabled = False
   
 mes = Month(FrmAuxiliarMovimientos.TxtFecha.Value)
 Año = Year(FrmAuxiliarMovimientos.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(FrmAuxiliarMovimientos.TxtFecha.Value) & "/" & Year(FrmAuxiliarMovimientos.TxtFecha.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 Fechas1 = Format(FechaIni, "yyyy/mm/dd")
 Fechas2 = Format(FechaFin, "yyyy/mm/dd")
FrmAuxiliarMovimientos.DtaConsulta.RecordSource = "SELECT CodCuentas, NombreCuenta, VoucherNo, DescripcionMovimiento, FacturaNo, ChequeNo, Clave, TCambio, Debito, Credito, FechaTransaccion, NPeriodo , NTransaccion, Fuente, FechaTasas, NumeroMovimiento From Transacciones WHERE  (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas2 & "', 102)) ORDER BY NumeroMovimiento"
'FrmAuxiliarMovimientos.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,  Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & "))ORDER BY Transacciones.NumeroMovimiento"
FrmAuxiliarMovimientos.DtaConsulta.Refresh
 
 If Not FrmAuxiliarMovimientos.DtaConsulta.Recordset.EOF Then
   NumeroTransaccion = Me.DBGCuentas.Columns(2).Text
'   Me.DtaSaldos.Recordset ("NumeroMovimiento")
  FrmAuxiliarMovimientos.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento,Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito,Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas,Transacciones.NumeroMovimiento , Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas2 & "', 102))AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ") ORDER BY Transacciones.NTransaccion"
  'FrmAuxiliarMovimientos.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,  Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion"
  FrmAuxiliarMovimientos.DtaTransacciones.Refresh
   If Not FrmAuxiliarMovimientos.DtaTransacciones.Recordset.EOF Then
     
    FrmAuxiliarMovimientos.TxtFecha.Value = FrmAuxiliarMovimientos.DtaTransacciones.Recordset("FechaTransaccion")
    FrmAuxiliarMovimientos.TxtPeriodo.Text = FrmAuxiliarMovimientos.DtaTransacciones.Recordset("Periodo")
    FrmAuxiliarMovimientos.TxtNTransacciones.Text = FrmAuxiliarMovimientos.DtaTransacciones.Recordset("NumeroMovimiento")
     NumeroTransaccion = FrmAuxiliarMovimientos.DtaTransacciones.Recordset("NumeroMovimiento")
    FrmAuxiliarMovimientos.TxtFuente.Text = FrmAuxiliarMovimientos.DtaTransacciones.Recordset("Fuente")
     '//////Sumo los Totales/////////////////////
    Debito = 0
    Credito = 0
    TotalDebito = 0
    TotalCredito = 0
      NumFecha1 = FrmAuxiliarMovimientos.TxtFecha.Value
      Fechas1 = Format(FrmAuxiliarMovimientos.TxtFecha.Value, "yyyy/mm/dd")
      NMovimiento = Val(FrmAuxiliarMovimientos.TxtNTransacciones)
     FrmAuxiliarMovimientos.DtaConsulta.RecordSource = "SELECT     FechaTransaccion, CodCuentas, NTransaccion, NumeroMovimiento, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, TCambio, Debito,Credito From Transacciones WHERE     (FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (NumeroMovimiento = " & NMovimiento & ")"
     'FrmAuxiliarMovimientos.DtaConsulta.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.CodCuentas, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, TCambio*Debito AS MDebito, TCambio*Credito AS MCredito, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito From Transacciones WHERE (((Transacciones.FechaTransaccion)=" & NumFecha1 & ") AND ((Transacciones.NumeroMovimiento)=" & NMovimiento & "))"
     FrmAuxiliarMovimientos.DtaConsulta.Refresh
      Do While Not FrmAuxiliarMovimientos.DtaConsulta.Recordset.EOF
       If Not IsNull(FrmAuxiliarMovimientos.DtaConsulta.Recordset("Debito")) Then
       Debito = FrmAuxiliarMovimientos.DtaConsulta.Recordset("Debito")
       End If
       If Not IsNull(Credito = FrmAuxiliarMovimientos.DtaConsulta.Recordset("Credito")) Then
        Credito = FrmAuxiliarMovimientos.DtaConsulta.Recordset("Credito")
       End If
       TotalDebito = TotalDebito + Debito
       TotalCredito = TotalCredito + Credito
      FrmAuxiliarMovimientos.DtaConsulta.Recordset.MoveNext
      Loop
   FrmAuxiliarMovimientos.TxtCredito.Text = Format(TotalCredito, "##,##0.00")
   FrmAuxiliarMovimientos.TxtDebito.Text = Format(TotalDebito, "##,##0.00")
   FrmAuxiliarMovimientos.TxtDiferencia.Text = Format(TotalDebito - TotalCredito, "##,##0.00")
        
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(0).Button = True
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(6).Button = True
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(6).Locked = True
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(0).Width = 1500
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(2).Width = 1000
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(3).Caption = "Descripcion"
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(4).Width = 1000
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(4).Button = True
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(5).Width = 1000
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(6).Width = 800
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).Width = 1200
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).NumberFormat = "##,##0.000000"
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(8).Width = 1200
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(8).NumberFormat = "##,##0.00"
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(9).Width = 1200
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(9).NumberFormat = "##,##0.00"
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(10).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(11).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(12).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(13).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(14).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(15).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(16).Visible = False
    FrmAuxiliarMovimientos.TxtFecha.Enabled = False
    FrmAuxiliarMovimientos.TxtPeriodo.Enabled = False
    FrmAuxiliarMovimientos.TxtFuente.Enabled = False
    FrmAuxiliarMovimientos.TxtNTransacciones.Enabled = False
    FrmAuxiliarMovimientos.DBGTransacciones.Enabled = True
    FrmAuxiliarMovimientos.CmbMoneda.Enabled = False
     End If
       
     FrmAuxiliarMovimientos.DBGTransacciones.Columns(0).Button = True
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(6).Button = True
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(6).Locked = True
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(0).Width = 1500
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(2).Width = 1000
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(3).Caption = "Descripcion"
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(4).Width = 1000
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(4).Button = True
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(5).Width = 1000
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(6).Width = 800
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).Width = 1200
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).NumberFormat = "##,##0.000000"
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(8).Width = 1200
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(8).NumberFormat = "##,##0.00"
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(9).Width = 1200
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(9).NumberFormat = "##,##0.00"
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(10).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(11).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(12).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(13).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(14).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(15).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(16).Visible = False
    FrmAuxiliarMovimientos.TxtFecha.Enabled = False
    FrmAuxiliarMovimientos.TxtPeriodo.Enabled = False
    FrmAuxiliarMovimientos.TxtFuente.Enabled = False
    FrmAuxiliarMovimientos.TxtNTransacciones.Enabled = False
    FrmAuxiliarMovimientos.DBGTransacciones.Enabled = True
        FrmAuxiliarMovimientos.CmbMoneda.Enabled = False
  End If






FrmAuxiliarMovimientos.Frame1.Enabled = False
FrmAuxiliarMovimientos.Show 1
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
'On Error GoTo TipoErrs
Dim Debito As Double, Credito As Double
Dim sqlconsulta As String
Dim c As Integer
Dim col As TrueOleDBGrid80.Column
Dim cols As TrueOleDBGrid80.Columns

Total1 = 0
QUIEN = "Nuevo"
Me.DtaCuentas.Refresh
Criterio = "CodCuentas='" & Me.DBCliente.Text & "'"
Me.DtaCuentas.Recordset.Find (Criterio)
If DtaCuentas.Recordset.EOF Then
 
   sqlconsulta = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento,VoucherNo, FacturaNo, ChequeNo, Transacciones.DescripcionMovimiento, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Debito*TCambio AS MDebito, TCambio*Credito AS MCredito From Transacciones Where (((Transacciones.CodCuentas) =  '0')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
'  Me.DtaSaldos.Refresh
        With rs
          .Close
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        
       Me.DBGCuentas.DataSource = rs
       
  Me.DBGCuentas.Columns(1).Caption = "Fecha"
  Me.DBGCuentas.Columns(1).Width = 1000
  Me.DBGCuentas.Columns(2).Caption = "No.Trans"
  Me.DBGCuentas.Columns(2).Width = 1000
  Me.DBGCuentas.Columns(3).Width = 1000
  Me.DBGCuentas.Columns(4).Width = 1000
  Me.DBGCuentas.Columns(5).Width = 1000
  Me.DBGCuentas.Columns(5).Caption = "Cheq/Rec"
  Me.DBGCuentas.Columns(6).Width = 3000
  Me.DBGCuentas.Columns(7).Width = 1000
  Me.DBGCuentas.Columns(7).NumberFormat = "##,##0.000000"
  Me.DBGCuentas.Columns(8).Visible = False
  Me.DBGCuentas.Columns(9).Visible = False
  Me.DBGCuentas.Columns(10).Width = 1200
  Me.DBGCuentas.Columns(10).Caption = "Debito"
  Me.DBGCuentas.Columns(11).Width = 1200
  Me.DBGCuentas.Columns(11).Caption = "Credito"
  Me.DBGCuentas.Columns(10).NumberFormat = "##,##0.00"
  Me.DBGCuentas.Columns(11).NumberFormat = "##,##0.00"
  Me.LblSaldoCuenta.Caption = "0.00"
  Me.DBGCuentas.Columns(0).Visible = False
  Me.CmbTipoMoneda.Enabled = True
  
Else
  Me.TxtDescripcion.Text = Me.DtaCuentas.Recordset("DescripcionCuentas")
'  Me.CmbTipoMoneda.Text = Me.DtaCuentas.Recordset("TipoMoneda")
  TipoCuenta = DtaCuentas.Recordset("TipoCuenta")
  If Not IsNull(DtaCuentas.Recordset("TipoMoneda")) Then
   Me.CmbTipoMoneda = DtaCuentas.Recordset("TipoMoneda")
  End If
 ' Me.LblSaldoCuenta.Caption = Format(DtaCuentas.Recordset.SaldoActual, "##,##0.00")

  '//////////Muestro los Saldos de las cuentas/////////////////////
  
 sqlconsulta = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento, VoucherNo, FacturaNo, ChequeNo,Transacciones.DescripcionMovimiento, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, ROUND(Debito*TCambio,3) AS MDebito, ROUND(TCambio*Credito,3) AS MCredito From Transacciones Where (((Transacciones.CodCuentas) =  '" & Me.DBCliente.Text & "')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
' Me.DtaSaldos.Refresh
        With rs
          .Close
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        
       Me.DBGCuentas.DataSource = rs
 
  Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento, Transacciones.DescripcionMovimiento, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, ROUND(Debito*TCambio,3) AS MDebito, ROUND(TCambio*Credito,3) AS MCredito From Transacciones Where (((Transacciones.CodCuentas) =  '" & Me.DBCliente.Text & "')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
  Me.DtaConsulta.Refresh
  
  Do While Not Me.DtaConsulta.Recordset.EOF
'   Me.CmbTipoMoneda.Enabled = False
   If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
    If Not IsNull(Me.DtaConsulta.Recordset("MDebito")) Then
     Debito = Format(Me.DtaConsulta.Recordset("MDebito"), "##,##0.00")
    End If
    If Not IsNull(Me.DtaConsulta.Recordset("MCredito")) Then
     Credito = Format(Me.DtaConsulta.Recordset("MCredito"), "##,##0.00")
    End If
    Total1 = Debito - Credito + Total1
    Debito = 0
    Credito = 0
   Else
    If Not IsNull(Me.DtaConsulta.Recordset("MDebito")) Then
     Debito = Format(Me.DtaConsulta.Recordset("MDebito"), "##,##0.00")
    End If
    If Not IsNull(Me.DtaConsulta.Recordset("MCredito")) Then
     Credito = Format(Me.DtaConsulta.Recordset("MCredito"), "##,##0.00")
    End If
    Total1 = Credito - Debito + Total1
    Debito = 0
    Credito = 0
   End If
   
   Me.DtaConsulta.Recordset.MoveNext
  Loop
  If TipoCuenta = "Bancos" Then
   Me.CmdAuditoria.Visible = False
   Me.CmdConciliacion.Visible = True
  Else
   Me.CmdAuditoria.Visible = True
   Me.CmdConciliacion.Visible = False
  End If
  Me.DBGCuentas.Columns(1).Caption = "Fecha"
  Me.DBGCuentas.Columns(1).Width = 1000
  Me.DBGCuentas.Columns(2).Caption = "No.Trans"
  Me.DBGCuentas.Columns(2).Width = 1000
  Me.DBGCuentas.Columns(3).Width = 1000
  Me.DBGCuentas.Columns(4).Width = 1000
  Me.DBGCuentas.Columns(5).Width = 1000
  Me.DBGCuentas.Columns(5).Caption = "Cheq/Rec"
  Me.DBGCuentas.Columns(6).Width = 3000
  Me.DBGCuentas.Columns(7).Width = 1000
  Me.DBGCuentas.Columns(7).NumberFormat = "##,##0.000000"
  Me.DBGCuentas.Columns(8).Visible = False
  Me.DBGCuentas.Columns(9).Visible = False
  Me.DBGCuentas.Columns(10).Width = 1200
  Me.DBGCuentas.Columns(10).Caption = "Debito"
  Me.DBGCuentas.Columns(11).Width = 1200
  Me.DBGCuentas.Columns(11).Caption = "Credito"
  Me.DBGCuentas.Columns(10).NumberFormat = "##,##0.00"
  Me.DBGCuentas.Columns(11).NumberFormat = "##,##0.00"
  
  Me.LblSaldoCuenta.Caption = Format(Total1, "##,##0.00")
  Me.DBGCuentas.Columns(0).Visible = False
 
 
 
 
 
End If
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub DBGCuentas_DblClick()
FrmAuxiliarMovimientos.TxtFecha.Value = Me.DBGCuentas.Columns(1).Text
NumeroTransaccion = Me.DBGCuentas.Columns(2).Text
  If Not Me.DBGCuentas.Columns(3).Text = "" Then
    Descripcion = Me.DBGCuentas.Columns(3).Text
   If Descripcion = "**********CANCELADO*************" Then
     MsgBox " Este movimiento esta Cancelado", vbCritical, "Sistema Contable"
     Exit Sub
   End If
  End If
   
    FrmAuxiliarMovimientos.Enabled = True
 FrmAuxiliarMovimientos.CmbMoneda.Enabled = False
   
 mes = Month(FrmAuxiliarMovimientos.TxtFecha.Value)
 Año = Year(FrmAuxiliarMovimientos.TxtFecha.Value)
 FechaIni = CDate("1/" & Month(FrmAuxiliarMovimientos.TxtFecha.Value) & "/" & Year(FrmAuxiliarMovimientos.TxtFecha.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
FrmAuxiliarMovimientos.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,  Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento From Transacciones WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & "))ORDER BY Transacciones.NumeroMovimiento"
FrmAuxiliarMovimientos.DtaConsulta.Refresh
 
 If Not FrmAuxiliarMovimientos.DtaConsulta.Recordset.EOF Then
   NumeroTransaccion = Me.DtaSaldos.Recordset("NumeroMovimiento")
  FrmAuxiliarMovimientos.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio,  Transacciones.Debito, Transacciones.Credito, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, Transacciones.NumeroMovimiento, Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.NumeroMovimiento)=" & NumeroTransaccion & ")) ORDER BY Transacciones.NTransaccion"
  FrmAuxiliarMovimientos.DtaTransacciones.Refresh
   If Not FrmAuxiliarMovimientos.DtaTransacciones.Recordset.EOF Then
     
    FrmAuxiliarMovimientos.TxtFecha.Value = FrmAuxiliarMovimientos.DtaTransacciones.Recordset("FechaTransaccion")
    FrmAuxiliarMovimientos.TxtPeriodo.Text = FrmAuxiliarMovimientos.DtaTransacciones.Recordset("Periodo")
    FrmAuxiliarMovimientos.TxtNTransacciones.Text = FrmAuxiliarMovimientos.DtaTransacciones.Recordset("NumeroMovimiento")
     NumeroTransaccion = FrmAuxiliarMovimientos.DtaTransacciones.Recordset("NumeroMovimiento")
    FrmAuxiliarMovimientos.TxtFuente.Text = FrmAuxiliarMovimientos.DtaTransacciones.Recordset("Fuente")
     '//////Sumo los Totales/////////////////////
    Debito = 0
    Credito = 0
    TotalDebito = 0
    TotalCredito = 0
      NumFecha1 = FrmAuxiliarMovimientos.TxtFecha.Value
      NMovimiento = Val(FrmAuxiliarMovimientos.TxtNTransacciones)
      Fechas1 = Format(FrmAuxiliarMovimientos.TxtFecha.Value, "yyyy/mm/dd")
      FrmAuxiliarMovimientos.DtaConsulta.RecordSource = "SELECT FechaTransaccion, CodCuentas, NTransaccion, NumeroMovimiento, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, TCambio, Debito,Credito From Transacciones WHERE     (FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (NumeroMovimiento = " & NMovimiento & ")"
'     FrmAuxiliarMovimientos.DtaConsulta.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.CodCuentas, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, [Transacciones]![TCambio]*[Transacciones]![Debito] AS MDebito, [Transacciones]![TCambio]*[Transacciones]![Credito] AS MCredito, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito From Transacciones WHERE (((Transacciones.FechaTransaccion)=" & NumFecha1 & ") AND ((Transacciones.NumeroMovimiento)=" & NMovimiento & "))"
     FrmAuxiliarMovimientos.DtaConsulta.Refresh
      Do While Not FrmAuxiliarMovimientos.DtaConsulta.Recordset.EOF
       If Not IsNull(FrmAuxiliarMovimientos.DtaConsulta.Recordset("Debito")) Then
       Debito = FrmAuxiliarMovimientos.DtaConsulta.Recordset("Debito")
       End If
       If Not IsNull(Credito = FrmAuxiliarMovimientos.DtaConsulta.Recordset("Credito")) Then
        Credito = FrmAuxiliarMovimientos.DtaConsulta.Recordset("Credito")
       End If
       TotalDebito = TotalDebito + Debito
       TotalCredito = TotalCredito + Credito
      FrmAuxiliarMovimientos.DtaConsulta.Recordset.MoveNext
      Loop
   FrmAuxiliarMovimientos.TxtCredito.Text = Format(TotalCredito, "##,##0.00")
   FrmAuxiliarMovimientos.TxtDebito.Text = Format(TotalDebito, "##,##0.00")
   FrmAuxiliarMovimientos.TxtDiferencia.Text = Format(TotalDebito - TotalCredito, "##,##0.00")
        
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(0).Button = True
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(6).Button = True
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(6).Locked = True
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(0).Width = 1500
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(2).Width = 1000
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(3).Caption = "Descripcion"
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(4).Width = 1000
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(4).Button = True
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(5).Width = 1000
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(6).Width = 800
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).Width = 1200
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).NumberFormat = "##,##0.000000"
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(8).Width = 1200
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(8).NumberFormat = "##,##0.00"
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(9).Width = 1200
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(9).NumberFormat = "##,##0.00"
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(10).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(11).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(12).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(13).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(14).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(15).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(16).Visible = False
    FrmAuxiliarMovimientos.TxtFecha.Enabled = False
    FrmAuxiliarMovimientos.TxtPeriodo.Enabled = False
    FrmAuxiliarMovimientos.TxtFuente.Enabled = False
    FrmAuxiliarMovimientos.TxtNTransacciones.Enabled = False
    FrmAuxiliarMovimientos.CmbMoneda.Enabled = False
    FrmAuxiliarMovimientos.DBGTransacciones.Enabled = True
     End If
       
     FrmAuxiliarMovimientos.DBGTransacciones.Columns(0).Button = True
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(5).Caption = "Cheq/Rec"
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(6).Button = True
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(6).Locked = True
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(0).Width = 1500
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(2).Width = 1000
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(3).Caption = "Descripcion"
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(4).Width = 1000
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(4).Button = True
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(5).Width = 1000
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(6).Width = 800
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).Width = 1200
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(7).NumberFormat = "##,##0.000000"
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(8).Width = 1200
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(8).NumberFormat = "##,##0.00"
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(9).Width = 1200
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(9).NumberFormat = "##,##0.00"
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(10).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(11).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(12).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(13).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(14).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(15).Visible = False
    FrmAuxiliarMovimientos.DBGTransacciones.Columns(16).Visible = False
    FrmAuxiliarMovimientos.TxtFecha.Enabled = False
    FrmAuxiliarMovimientos.TxtPeriodo.Enabled = False
    FrmAuxiliarMovimientos.TxtFuente.Enabled = False
    FrmAuxiliarMovimientos.TxtNTransacciones.Enabled = False
    FrmAuxiliarMovimientos.CmbMoneda.Enabled = False
    FrmAuxiliarMovimientos.DBGTransacciones.Enabled = True
  End If






FrmAuxiliarMovimientos.Frame1.Enabled = False
FrmAuxiliarMovimientos.Show 1
End Sub

Private Sub DBGCuentas_FilterChange()
'On Error GoTo TipoErrs:
'Dim Filtro As String
'Set cols = DBGCuentas.Columns
'Dim c As Integer
'c = DBGCuentas.col
'DBGCuentas.HoldFields
'Filtro = getFilter()
'DtaSaldos.Recordset.Filter = Filtro
'DBGCuentas.col = c
'DBGCuentas.EditActive = True

    Dim col As TrueOleDBGrid80.Column
    Dim cols As TrueOleDBGrid80.Columns
    
    'On Error GoTo errHandler
    On Error Resume Next
    Set cols = Me.DBGCuentas.Columns
    Dim c As Integer
    
    c = DBGCuentas.col
    DBGCuentas.HoldFields
    SQL = rs.Filter
    rs.Filter = getFilter(col, cols)
    DBGCuentas.col = c
    DBGCuentas.EditActive = True
Exit Sub
TipoErrs:
 MsgBox err.Description
End Sub
Private Function getFilter(col As TrueOleDBGrid80.Column, cols As TrueOleDBGrid80.Columns) As String
'Creates the SQL statement in adodc1.recordset.filter
'and only filters text currently. It must be modified to
'filter other data types.
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
getFilter = tmp

End Function

'Function getFilter() As String
'
'Dim tmp As String
'Dim n As Integer
'For Each col In cols
'If Trim(col.FilterText) <> "" Then
'n = n + 1
'If n > 1 Then
'tmp = tmp & " AND "
'End If
'tmp = tmp & col.DataField & " LIKE '" & col.FilterText & "*'"
'End If
'Next col
'
'getFilter = tmp
'End Function

Private Sub Form_Activate()
' Me.DtaSaldos.Refresh

 Me.DtaCuentas.Refresh
  Me.DBGCuentas.Columns(1).Caption = "Fecha"
  Me.DBGCuentas.Columns(1).Width = 1000
  Me.DBGCuentas.Columns(2).Caption = "No.Trans"
  Me.DBGCuentas.Columns(2).Width = 1000
  Me.DBGCuentas.Columns(3).Width = 1000
  Me.DBGCuentas.Columns(4).Width = 1000
  Me.DBGCuentas.Columns(5).Width = 1000
  Me.DBGCuentas.Columns(5).Caption = "Cheq/Rec"
  Me.DBGCuentas.Columns(6).Width = 3000
  Me.DBGCuentas.Columns(7).Width = 1000
  Me.DBGCuentas.Columns(7).NumberFormat = "##,##0.000000"
  Me.DBGCuentas.Columns(8).Visible = False
  Me.DBGCuentas.Columns(9).Visible = False
  Me.DBGCuentas.Columns(10).Width = 1200
  Me.DBGCuentas.Columns(10).Caption = "Debito"
  Me.DBGCuentas.Columns(11).Width = 1200
  Me.DBGCuentas.Columns(11).Caption = "Credito"
  Me.DBGCuentas.Columns(10).NumberFormat = "##,##0.00"
  Me.DBGCuentas.Columns(11).NumberFormat = "##,##0.00"
  Me.DBGCuentas.Columns(0).Visible = False
End Sub

Private Sub Form_Load()

 MDIPrimero.Skin1.ApplySkin hWnd
 Me.DBGCuentas.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DBGCuentas.OddRowStyle.BackColor = &H80000005
 Me.DBGCuentas.AlternatingRowStyle = True
 'AZUL Y BLANCO COMO LA PATRIA
With Me.DtaConsulta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With
With Me.DtaCuentas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Cuentas"
End With
With Me.DtaSaldos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

If cnx.State = adStateClosed Then
    cnx.ConnectionString = Conexion
    cnx.Open
End If

 sqlconsulta = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento, VoucherNo, FacturaNo, ChequeNo,Transacciones.DescripcionMovimiento, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Debito*TCambio AS MDebito, TCambio*Credito AS MCredito From Transacciones Where (((Transacciones.CodCuentas) =  '" & Me.DBCliente.Text & "')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
' Me.DtaSaldos.Refresh

        With rs
          .CursorLocation = adUseClient
          .Open sqlconsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
        
        
       Me.DBGCuentas.DataSource = rs

  Me.DBGCuentas.Columns(2).Caption = "No.Trans"
  Me.DBGCuentas.Columns(2).Width = 1000
  Me.DBGCuentas.Columns(3).Width = 500
  Me.DBGCuentas.Columns(4).Width = 500
  Me.DBGCuentas.Columns(5).Width = 500
  Me.DBGCuentas.Columns(6).Width = 3000
  Me.DBGCuentas.Columns(8).Visible = False
  Me.DBGCuentas.Columns(9).Visible = False
  Me.DBGCuentas.Columns(10).Caption = "Debito"
  Me.DBGCuentas.Columns(11).Caption = "Credito"
  Me.DBGCuentas.Columns(10).NumberFormat = "##,##0.00"
  Me.DBGCuentas.Columns(11).NumberFormat = "##,##0.00"
  Me.DBGCuentas.Columns(7).NumberFormat = "##,##0.000000"
  Me.DBGCuentas.Columns(0).Visible = False


End Sub

Private Sub SmartButton1_Click()

End Sub
