VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmTransferencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencia de Saldos"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8580
   Icon            =   "FrmTranfiere.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   8580
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc DtaCuentaOrigen 
      Height          =   375
      Left            =   240
      Top             =   4320
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
      Caption         =   "DtaCuentaOrigen"
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
      Top             =   3840
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
   Begin MSAdodcLib.Adodc DtaTranscciones 
      Height          =   375
      Left            =   240
      Top             =   3360
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
      Caption         =   "DtaTranscciones"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   7200
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton CmdProcesar 
         Caption         =   "Procesar"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Txtorigen 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox TxtDestino 
         Height          =   285
         Left            =   6000
         TabIndex        =   3
         Top             =   240
         Width           =   1695
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
         Left            =   3000
         Picture         =   "FrmTranfiere.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton CmdBuscaCuetaDestino 
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
         Left            =   7800
         Picture         =   "FrmTranfiere.frx":0A18
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmTranfiere.frx":0B66
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   4680
         OleObjectBlob   =   "FrmTranfiere.frx":0BDE
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   3720
         OleObjectBlob   =   "FrmTranfiere.frx":0C5A
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin XtremeSuiteControls.ProgressBar BarCalcular 
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   5655
         _Version        =   786432
         _ExtentX        =   9975
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   14737632
         Scrolling       =   1
         Appearance      =   6
      End
   End
End
Attribute VB_Name = "FrmTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBuscaCuenta_Click()
QueProducto = "Transferencia1"
FrmConsulta.Show 1
End Sub

Private Sub CmdBuscaCuetaDestino_Click()
QueProducto = "Transferencia2"
FrmConsulta.Show 1
End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub CmdProcesar_Click()
On Error GoTo TipoErrs
Dim MonedaOrigen As String, MonedaDestino As String, KeyOrigen As String, KeyDestino As String
Dim i As Integer, CantRegistros As Integer
Me.Frame1.Enabled = False


Me.DtaCuentaOrigen.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.CodGrupo, Cuentas.SaldoActual, Cuentas.TipoMoneda, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo From Cuentas Where (((Cuentas.CodCuentas) = '" & Me.Txtorigen.Text & "'))"
Me.DtaCuentaOrigen.Refresh
KeyOrigen = Mid(Me.DtaCuentaOrigen.Recordset("KeyGrupo"), 1, 1)
MonedaOrigen = Me.DtaCuentaOrigen.Recordset("TipoMoneda")

Me.DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.CodGrupo, Cuentas.SaldoActual, Cuentas.TipoMoneda, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo From Cuentas Where (((Cuentas.CodCuentas) = '" & Me.TxtDestino.Text & "'))"
Me.DtaConsulta.Refresh
MonedaDestino = Me.DtaConsulta.Recordset("TipoMoneda")
KeyDestino = Mid(Me.DtaConsulta.Recordset("KeyGrupo"), 1, 1)
If MonedaDestino = MonedaOrigen Then
  If KeyOrigen = KeyDestino Then
    Me.DtaTranscciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NumeroMovimiento, Transacciones.NombreCuenta, Transacciones.DescripcionMovimiento, Cuentas.TipoMoneda, Transacciones.NPeriodo, Transacciones.FechaTransaccion FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas Where (((Transacciones.CodCuentas) = '" & Me.Txtorigen.Text & "')) ORDER BY Transacciones.FechaTransaccion"
    Me.DtaTranscciones.Refresh
    Me.DtaTranscciones.Recordset.MoveLast
    CantRegistros = Me.DtaTranscciones.Recordset.RecordCount
    Me.DtaTranscciones.Recordset.MoveFirst
    Me.BarCalcular.Visible = True
    With BarCalcular
     .Min = 0
     .Max = CantRegistros
     .Value = 0
     i = 1
     Do While Not Me.DtaTranscciones.Recordset.EOF
      .Value = i
'      Me.DtaTranscciones.Recordset.Edit
       Me.DtaTranscciones.Recordset("CodCuentas") = Me.TxtDestino.Text
      Me.DtaTranscciones.Recordset.Update
      Me.DtaTranscciones.Recordset.MoveNext
      i = i + 1
     Loop
    End With
   Me.DtaCuentaOrigen.Recordset.Delete
    Unload Me
  Else
   MsgBox " No Coincide el tipo de Cuenta", vbCritical, "Sistema contable"
   Me.Frame1.Enabled = True
   Exit Sub
  End If
Else
  MsgBox "No Coincide el tipo de monedas de las cuentas", vbCritical, "Sistema Contable"
  Me.Frame1.Enabled = True
  Exit Sub
End If
Exit Sub
TipoErrs:
 MsgBox err.Description
End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd
'Me.DtaTranscciones.DatabaseName = Ruta
Me.DtaTranscciones.ConnectionString = Conexion

'Me.DtaConsulta.DatabaseName = Ruta
Me.DtaConsulta.ConnectionString = Conexion

'Me.DtaCuentaOrigen.DatabaseName = Ruta
Me.DtaCuentaOrigen.ConnectionString = Conexion

End Sub

