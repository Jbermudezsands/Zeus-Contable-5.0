VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmCalcularDepreciacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calcular Depreciacion"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "FrmCalcularDepreciacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ProgressBar BarCalcular 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   4335
      _Version        =   786432
      _ExtentX        =   7646
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   615
      Left            =   120
      OleObjectBlob   =   "FrmCalcularDepreciacion.frx":57E2
      TabIndex        =   7
      Top             =   240
      Width           =   4455
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton CmdCalcular 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "FrmCalcularDepreciacion.frx":599A
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmCalcularDepreciacion.frx":5A3E
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo DCmbCodigo 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc AdoConsulta 
      Height          =   330
      Left            =   5040
      Top             =   2520
      Width           =   2415
      _ExtentX        =   4260
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
   Begin MSAdodcLib.Adodc AdoTasas 
      Height          =   330
      Left            =   5040
      Top             =   2205
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "AdoTasas"
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
   Begin MSAdodcLib.Adodc AdoIndice 
      Height          =   330
      Left            =   5040
      Top             =   1875
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "AdoIndice"
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
   Begin MSAdodcLib.Adodc AdoPeriodos 
      Height          =   330
      Left            =   5040
      Top             =   1560
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "AdoPeriodos"
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
   Begin MSAdodcLib.Adodc AdoActivoFijo 
      Height          =   330
      Left            =   5040
      Top             =   1245
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "AdoActivoFijo"
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
   Begin MSAdodcLib.Adodc AdoTransacciones 
      Height          =   330
      Left            =   5040
      Top             =   915
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "AdoTransacciones"
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
   Begin MSAdodcLib.Adodc AdoCuentas 
      Height          =   330
      Left            =   5040
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
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
   Begin MSComCtl2.DTPicker TxtFecha 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   503
      _Version        =   393216
      Format          =   17104897
      CurrentDate     =   38008
   End
   Begin VB.Label LblNombre 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   4215
   End
End
Attribute VB_Name = "FrmCalcularDepreciacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub CmdCalcular_Click()
Dim CanRegistros As Integer, i As Integer
Dim ValorOriginal As Double, ValorRescate As Double, VidaEstimada As Double
Dim Depreciacion As Double, TotalDepreciacion As Double, NumFecha As Long
Dim CuentaDepreciacion As String, CuentaGasto As String, Tasas As Double
Dim TipoCuentaGastos As String, TipoCuentaDepreciacion As String
Dim Debito As Double, Credito As Double, CuentaValorOriginal As String
On Error GoTo TipoErrs
NumeroTransaccion = NumeroTransaccion + 1
 'If Me.DCmbCodigo.Text = "" Then
  'MsgBox "Se necesita la cuenta de Depreciacion", vbCritical, "sistema Contable"
  'Exit Sub
 'End If
 NumFecha = Me.txtfecha.Value
 Me.AdoTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas From Tasas Where (((Tasas.FechaTasas) = " & NumFecha & "))"
 Me.AdoTasas.Refresh
 If Me.AdoTasas.Recordset.EOF Then
  MsgBox "No existe la Tasa de Cambio para la fecha del Movimiento", vbCritical, "sistema Contable"
  Exit Sub
 Else
  Tasas = Me.AdoTasas.Recordset!MontoCordobas
 End If
 
 
 Me.AdoActivoFijo.RecordSource = "SELECT ActivoFijo.CuentaGastos,ActivoFijo.CuentaDepreciacion,Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, ActivoFijo.ValorOriginal, ActivoFijo.FechaUltimaDepre, ActivoFijo.ValorEstimadoMeses, ActivoFijo.ValorRescate, ActivoFijo.DepreciacionAcumulada FROM Cuentas INNER JOIN ActivoFijo ON Cuentas.CodCuentas = ActivoFijo.CodCuenta ORDER BY Cuentas.CodCuentas  "
 Me.AdoActivoFijo.Refresh
 Me.AdoActivoFijo.Recordset.MoveLast
 CanRegistros = Me.AdoActivoFijo.Recordset.RecordCount
 Me.AdoActivoFijo.Recordset.MoveFirst
 MsgBox ("Se Procesarán " & CanRegistros & " Activo Fijos")
 With BarCalcular
 .Min = 0
 .Max = CanRegistros
 .Value = 0
 i = 1
            'Valido que no hayan duplicados de indices de transacciones JP
            AdoConsulta.RecordSource = "Select * from IndiceTransaccion where FechaTransaccion='" & Format(txtfecha, "yyyymmdd") & "' and NumeroMovimiento=" & Str(NumeroTransaccion)
            AdoConsulta.Refresh
            If Not AdoConsulta.Recordset.EOF Then
                MsgBox "Ya se ha hecho esta Transacción Anteriormente", vbInformation
                Exit Sub
            End If
   'Agrego el indice
          Me.AdoIndice.Recordset.AddNew
          Me.AdoIndice.Recordset!FechaTransaccion = Me.txtfecha.Value
          Me.AdoIndice.Recordset!NumeroMovimiento = NumeroTransaccion
          Me.AdoIndice.Recordset!DescripcionMovimiento = "Calculo Automatico Depreciacion"
          Me.AdoIndice.Recordset!Fuente = "DEPRECIACION"
          Me.AdoIndice.Recordset!NPeriodo = NumeroPeriodo
           Me.AdoIndice.Recordset!TipoMoneda = "Córdobas"
          Me.AdoIndice.Recordset.Update
  
 Do While Not Me.AdoActivoFijo.Recordset.EOF
      .Value = i
    CuentaValorOriginal = Me.AdoActivoFijo.Recordset!CodCuentas
    CuentaDepreciacion = Me.AdoActivoFijo.Recordset!CuentaDepreciacion
    CuentaGasto = Me.AdoActivoFijo.Recordset!CuentaGastos
   'Busco si existe la Cuenta
   Me.AdoConsulta.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda From Cuentas WHERE (((Cuentas.CodCuentas)='" & CuentaGasto & "'))"
   Me.AdoConsulta.Refresh
'/////Busco la Cuenta de Gastos////////////////////////////
  If Not AdoConsulta.Recordset.EOF Then
   TipoCuentaGastos = Me.AdoConsulta.Recordset!TipoMoneda
   Me.AdoConsulta.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda From Cuentas WHERE (((Cuentas.CodCuentas)='" & CuentaDepreciacion & "'))"
   Me.AdoConsulta.Refresh
'///////////Busco la cuenta de la Depreciacion//////////
   If Not AdoConsulta.Recordset.EOF Then
    TipoCuentaDepreciacion = Me.AdoConsulta.Recordset!TipoMoneda
 '///////////////Busco el Saldo para el Valor Original//////////////////////////////////
  Me.AdoConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Transacciones.Debito*Transacciones.TCambio) AS MDebito, Sum(Transacciones.TCambio*Transacciones.Credito) AS MCredito From Transacciones GROUP BY Transacciones.CodCuentas Having (((Transacciones.CodCuentas) = '" & CuentaValorOriginal & "'))"
  Me.AdoConsulta.Refresh
  If Not Me.AdoConsulta.Recordset.EOF Then
      Debito = Me.AdoConsulta.Recordset!MDebito
      Credito = Me.AdoConsulta.Recordset!MCredito
      ValorOriginal = Debito - Credito
  
  End If
    

    ValorRescate = Me.AdoActivoFijo.Recordset!ValorRescate
    VidaEstimada = Me.AdoActivoFijo.Recordset!ValorEstimadoMeses
    
    
    
    
    If Not VidaEstimada = 0 Then
      Depreciacion = (ValorOriginal - ValorRescate) / VidaEstimada
    Else
      Depreciacion = 0
    End If
    CodigoCuenta = Me.AdoActivoFijo.Recordset!CodCuentas
    
    
    'aGREO LA TRANSACCION
    'No agrega la en la transacción el valor del campo NTransacción que está como parte de la llave y no puede quedar nula JP
   Me.AdoTransacciones.Recordset.AddNew
   Me.AdoTransacciones.Recordset!CodCuentas = CuentaDepreciacion
   Me.AdoTransacciones.Recordset!FechaTransaccion = Me.txtfecha.Value
   Me.AdoTransacciones.Recordset!NPeriodo = NumeroPeriodo
   Me.AdoTransacciones.Recordset!NumeroMovimiento = NumeroTransaccion
   Me.AdoTransacciones.Recordset!NombreCuenta = "Calculo Automatico Depreciacion"
   Me.AdoTransacciones.Recordset!DescripcionMovimiento = "Movimiento de Depreciacion"
   Me.AdoTransacciones.Recordset!Clave = "Credito"
   Me.AdoTransacciones.Recordset!Credito = Depreciacion
   Me.AdoTransacciones.Recordset!Fuente = "DEPRECIACION"
   If TipoCuentaDepreciacion = "Córdobas" Then
     Me.AdoTransacciones.Recordset!TCambio = 1
   Else
      Me.AdoTransacciones.Recordset!TCambio = 1 / Tasas
   End If
  Me.AdoTransacciones.Recordset.Update
  
     Me.AdoTransacciones.Recordset.AddNew
   Me.AdoTransacciones.Recordset!CodCuentas = CuentaGasto
   Me.AdoTransacciones.Recordset!FechaTransaccion = Me.txtfecha.Value
   Me.AdoTransacciones.Recordset!NPeriodo = NumeroPeriodo
   Me.AdoTransacciones.Recordset!NumeroMovimiento = NumeroTransaccion
   Me.AdoTransacciones.Recordset!NombreCuenta = "Calculo Automatico Depreciacion"
   Me.AdoTransacciones.Recordset!DescripcionMovimiento = "Movimiento de Depreciacion"
   Me.AdoTransacciones.Recordset!Clave = "Debito"
   Me.AdoTransacciones.Recordset!Debito = Depreciacion
   Me.AdoTransacciones.Recordset!Fuente = "DEPRECIACION"
   If TipoCuentaGastos = "Córdobas" Then
     Me.AdoTransacciones.Recordset!TCambio = 1
   Else
      Me.AdoTransacciones.Recordset!TCambio = 1 / Tasas
   End If
  Me.AdoTransacciones.Recordset.Update
  
  'Agrego activo fijo
   'Me.AdoActivoFijo.Recordset.Edit
     'Me.AdoActivoFijo.Recordset.DepreciacionAcumulada = Val(Me.AdoActivoFijo.Recordset.DepreciacionAcumulada) + Depreciacion
   
   'Me.AdoActivoFijo.Recordset.Update
    
    TotalDepreciacion = Depreciacion + TotalDepreciacion
   End If
  End If
  Me.AdoActivoFijo.Recordset.MoveNext
  
  
  i = i + 1
  Loop
  
'   Me.AdoTransacciones.Recordset.AddNew
'   Me.AdoTransacciones.Recordset.CodCuentas = Me.DCmbCodigo.Text
'   Me.AdoTransacciones.Recordset("FechaTransaccion") = Me.TxtFecha.Value
'   Me.AdoTransacciones.Recordset.NPeriodo = NumeroPeriodo
'   Me.AdoTransacciones.Recordset("NumeroMovimiento") = NumeroTransaccion
'   Me.AdoTransacciones.Recordset.NombreCuenta = "Calculo Automatico Depreciacion"
'   Me.AdoTransacciones.Recordset.DescripcionMovimiento = "Movimiento de Depreciacion"
'   Me.AdoTransacciones.Recordset.Clave = "Debito"
'   Me.AdoTransacciones.Recordset.Debito = TotalDepreciacion
'   Me.AdoTransacciones.Recordset("Fuente") = "DEPRECIACION"
'   Me.AdoTransacciones.Recordset.Tcambio = 1
'  Me.AdoTransacciones.Recordset.Update
 
 'edito periodos
 Mes = Month(Me.txtfecha.Value)
 Año = Year(Me.txtfecha.Value)
 FechaIni = CDate("1/" & Month(Me.txtfecha.Value) & "/" & Year(Me.txtfecha.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DCmbCodigo.Enabled = True
 Me.AdoConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
 Me.AdoConsulta.Refresh
 If Not AdoConsulta.Recordset.EOF Then
  'AdoConsulta.Recordset.Edit
   AdoConsulta.Recordset!NTransacciones = AdoConsulta.Recordset!NTransacciones + 1
  AdoConsulta.Recordset.Update
 End If
 
 End With
 
 MsgBox "El Proceso ha Finalizado Correctamente", vbInformation, "Sistema Contable"
Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub DCmbCodigo_Change()
On Error GoTo TipoErrs
Criterio = "CodCuentas='" & Me.DCmbCodigo.Text & "'"
Me.AdoCuentas.Recordset.Find (Criterio)
If Not AdoCuentas.Recordset.EOF Then
 If Not Me.AdoCuentas.Recordset.EOF Then
   Me.LblNombre.Caption = Me.AdoCuentas.Recordset!DescripcionCuentas
   Me.CmdCalcular.Enabled = True
 End If
End If
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub Form_Activate()
'Me.TxtFecha.Value = Format(FechaSistema, "dd/mm/yyyy")

Me.AdoCuentas.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.TipoCuenta, Cuentas.DescripcionCuentas From Cuentas Where (((Cuentas.TipoCuenta) = 'Gastos')) ORDER BY Cuentas.CodCuentas"
Me.AdoCuentas.Refresh
'Me.DCmbCodigo.ListField = "CodCuentas"
End Sub

Private Sub Form_Load()
On Error GoTo TipoErrs
MDIPrimero.Skin1.ApplySkin hWnd
With Me.AdoCuentas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from Cuentas"
   .Refresh
End With

With Me.AdoActivoFijo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from ActivoFijo"
   .Refresh
End With

With Me.AdoConsulta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.AdoIndice
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from IndiceTransaccion"
   .Refresh
End With

With Me.AdoPeriodos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from Periodos"
   .Refresh
End With

With Me.AdoTasas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from Tasas"
   .Refresh
End With

With Me.AdoTransacciones
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from Transacciones"
   .Refresh
End With



Me.txtfecha.Value = Format(Now, "dd/mm/yyyy")
AdoCuentas.ConnectionString = Conexion
Me.AdoCuentas.RecordSource = "SELECT top 10 Cuentas.CodCuentas, Cuentas.TipoCuenta, Cuentas.DescripcionCuentas From Cuentas Where (((Cuentas.TipoCuenta) = 'Cuentas de Gastos')) ORDER BY Cuentas.CodCuentas"
Me.AdoCuentas.Refresh
LlenarDataCombos AdoCuentas, DCmbCodigo, "CodCuentas", "CodCuentas"
'Me.DCmbCodigo.ListField = "CodCuentas"

Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub TxtFecha_Change()
On Error GoTo TipoErrs
   Me.CmdCalcular.Enabled = True
 'Me.DBGTransacciones.Enabled = True
 Mes = Month(Me.txtfecha.Value)
 Año = Year(Me.txtfecha.Value)
 FechaIni = CDate("1/" & Month(Me.txtfecha.Value) & "/" & Year(Me.txtfecha.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DCmbCodigo.Enabled = True
 Me.AdoConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
 Me.AdoConsulta.Refresh
 If Not AdoConsulta.Recordset.EOF Then
  NumeroPeriodo = AdoConsulta.Recordset!NPeriodo
  NumeroTransaccion = AdoConsulta.Recordset!NTransacciones
  EstadoPeriodo = AdoConsulta.Recordset!EstadoPeriodo
  If EstadoPeriodo = "B" Then
   MsgBox "El Periodo Esta Bloqueado", vbCritical, "Sistema Contable"
   Me.txtfecha.SetFocus
   
   Exit Sub
  ElseIf EstadoPeriodo = "C" Then
  MsgBox "El Periodo esta Cerrado", vbCritical, "Sistema Contable"
  Me.txtfecha.SetFocus
  txtfecha.Enabled = True
  
  Exit Sub
  Else
   'Me.DBGTransacciones.Enabled = True
  End If
 Else
   MsgBox "La Fecha esta fuera del Rango de Periodos", vbCritical, "Sistema Contable"
   'Me.DBGTransacciones.Enabled = False
   txtfecha.Enabled = True
   
   Exit Sub
 End If
 
 '///////Verifico si esta registrada la fecha de la tasa//////



Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub TxtFecha_GotFocus()
On Error GoTo TipoErrs
 'Me.DBGTransacciones.Enabled = True
 Mes = Month(Me.txtfecha.Value)
 Año = Year(Me.txtfecha.Value)
 FechaIni = CDate("1/" & Month(Me.txtfecha.Value) & "/" & Year(Me.txtfecha.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DCmbCodigo.Enabled = True
 Me.AdoConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
 Me.AdoConsulta.Refresh
 If Not AdoConsulta.Recordset.EOF Then
  NumeroPeriodo = AdoConsulta.Recordset!NPeriodo
  NumeroTransaccion = AdoConsulta.Recordset!NTransacciones
  EstadoPeriodo = AdoConsulta.Recordset!EstadoPeriodo
  If EstadoPeriodo = "B" Then
   MsgBox "El Periodo Esta Bloqueado", vbCritical, "Sistema Contable"
   Me.txtfecha.SetFocus
   
   Exit Sub
  ElseIf EstadoPeriodo = "C" Then
  MsgBox "El Periodo esta Cerrado", vbCritical, "Sistema Contable"
  Me.txtfecha.SetFocus
  txtfecha.Enabled = True
  
  Exit Sub
  Else
   'Me.DBGTransacciones.Enabled = True
  End If
 Else
   MsgBox "La Fecha esta fuera del Rango de Periodos", vbCritical, "Sistema Contable"
   'Me.DBGTransacciones.Enabled = False
   txtfecha.Enabled = True
   
   Exit Sub
 End If
 
 '///////Verifico si esta registrada la fecha de la tasa//////



Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub TxtFecha_LostFocus()
On Error GoTo TipoErrs
 'Me.DBGTransacciones.Enabled = True
 Mes = Month(Me.txtfecha.Value)
 Año = Year(Me.txtfecha.Value)
 FechaIni = CDate("1/" & Month(Me.txtfecha.Value) & "/" & Year(Me.txtfecha.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DCmbCodigo.Enabled = True
 Me.AdoConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
 Me.AdoConsulta.Refresh
 If Not AdoConsulta.Recordset.EOF Then
  NumeroPeriodo = AdoConsulta.Recordset!NPeriodo
  NumeroTransaccion = AdoConsulta.Recordset!NTransacciones
  EstadoPeriodo = AdoConsulta.Recordset!EstadoPeriodo
  If EstadoPeriodo = "B" Then
   MsgBox "El Periodo Esta Bloqueado", vbCritical, "Sistema Contable"
   Me.txtfecha.SetFocus
   
   Exit Sub
  ElseIf EstadoPeriodo = "C" Then
  MsgBox "El Periodo esta Cerrado", vbCritical, "Sistema Contable"
  Me.txtfecha.SetFocus
  txtfecha.Enabled = True
  
  Exit Sub
  Else
   'Me.DBGTransacciones.Enabled = True
  End If
 Else
   MsgBox "La Fecha esta fuera del Rango de Periodos", vbCritical, "Sistema Contable"
   'Me.DBGTransacciones.Enabled = False
   txtfecha.Enabled = True
   
   Exit Sub
 End If
 
 '///////Verifico si esta registrada la fecha de la tasa//////



Exit Sub
TipoErrs:
 ControlErrores
End Sub
