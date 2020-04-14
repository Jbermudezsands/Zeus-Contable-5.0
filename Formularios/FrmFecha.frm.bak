VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Begin VB.Form FrmFecha 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbMoneda 
      Height          =   315
      ItemData        =   "FrmFecha.frx":0000
      Left            =   2040
      List            =   "FrmFecha.frx":000D
      TabIndex        =   5
      Text            =   "Córdobas"
      Top             =   840
      Width           =   2295
   End
   Begin VB.Data DtaConsulta 
      Caption         =   "DtaConsulta"
      Connect         =   ";DATABASENAME="" + Ruta + "";UID=Administrador;PWD=DFID"
      DatabaseName    =   "D:\DFID\dfid.bak"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Data DtaTasas 
      Caption         =   "DtaTasas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   4995
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCELAR"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Format          =   21364737
      CurrentDate     =   38430
   End
   Begin vbskfree.Skinner Skinner1 
      Left            =   960
      Top             =   2040
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin VB.Label Label2 
      Caption         =   "Moneda Movimiento:"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Transaccion"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "FrmFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub CmdOK_Click()
On Error GoTo TipoErrs
 FrmAuxiliarMovimientos.CmbMoneda.Text = Me.CmbMoneda.Text
 FrmAuxiliarMovimientos.TxtFuente = "Conciliacion"
 
 FrmAuxiliarMovimientos.Frame1.Enabled = False
 FrmAuxiliarMovimientos.DBGTransacciones.Enabled = True
 Mes = Month(Me.DTPicker1.Value)
 Año = Year(Me.DTPicker1.Value)
 FechaIni = CDate("1/" & Month(Me.DTPicker1.Value) & "/" & Year(Me.DTPicker1.Value))
 FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
 Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
 FrmAuxiliarMovimientos.TxtPeriodo.Text = DtaConsulta.Recordset("Periodo")
  NumeroPeriodo = DtaConsulta.Recordset("NPeriodo")
  NumeroTransaccion = DtaConsulta.Recordset("NTransacciones")
  EstadoPeriodo = DtaConsulta.Recordset("EstadoPeriodo")
  If EstadoPeriodo = "A" Then
      '  TipoCuenta = Me.DtaProductos.Recordset("TipoCuenta")
    CodigoCuenta = FrmConciliacion.DBCliente.Text
     FrmAuxiliarMovimientos.DBGTransacciones.Columns(0) = FrmConciliacion.DBCliente.Text
     Criterio = "CodCuentas='" & FrmAuxiliarMovimientos.DBGTransacciones.Columns(0).Text & "'"
       FrmAuxiliarMovimientos.DtaCuentas.Recordset.Find (Criterio)
       If Not FrmAuxiliarMovimientos.DtaCuentas.Recordset.EOF Then
         FrmAuxiliarMovimientos.CmbMoneda.Enabled = False
         Mes = Month(FrmAuxiliarMovimientos.TxtFecha.Value)
         Año = Year(FrmAuxiliarMovimientos.TxtFecha.Value)
         FechaIni = CDate("1/" & Month(FrmAuxiliarMovimientos.TxtFecha.Value) & "/" & Year(FrmAuxiliarMovimientos.TxtFecha.Value))
         FechaFin = DateSerial(Año, Mes + 1, 1 - 1)
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
          FrmAuxiliarMovimientos.DtaIndice.Recordset("FechaTransaccion") = FrmAuxiliarMovimientos.TxtFecha.Value
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
       
        
   'TipoCuenta = Me.DtaProductos.Recordset("TipoCuenta")
   CodigoCuenta = FrmConciliacion.DBCliente.Text
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
  
  
  ElseIf EstadoPeriodo = "B" Then
   MsgBox "El Periodo Esta Bloqueado", vbCritical, "Sistema Contable"

   Exit Sub
  ElseIf EstadoPeriodo = "C" Then
  MsgBox "El Periodo esta Cerrado", vbCritical, "Sistema Contable"
  Exit Sub
  Else
   FrmAuxiliarMovimientos.DBGTransacciones.Enabled = True
  If Not CodigoUsuario = 0 Then
  FrmAuxiliarMovimientos.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Transacciones'))"
   FrmAuxiliarMovimientos.DtaNacceso.Refresh
   If FrmAuxiliarMovimientos.DtaNacceso.Recordset.EOF Then
    FrmAuxiliarMovimientos.DBGTransacciones.Enabled = False
   Else
     FrmAuxiliarMovimientos.DBGTransacciones.Enabled = True
   End If
      
  End If
  End If
 Else
   MsgBox "La Fecha esta fuera del Rango de Periodos", vbCritical, "Sistema Contable"

   Exit Sub
 End If
 
 '///////Verifico si esta registrada la fecha de la tasa//////

NumFecha = Me.DTPicker1.Value
Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) = " & NumFecha & "))ORDER BY Tasas.FechaTasas"
Me.DtaTasas.Refresh

If Not DtaTasas.Recordset.EOF Then
Fecha = Format(DtaTasas.Recordset("FechaTasas"), "dd/mm/yyyy")
   
    Encontrado = True
    Cambio = DtaTasas.Recordset("MontoCordobas")
    MDIPrimero.StatusBar2.Panels(2) = "Tasa Dolar: " & Format(Cambio, "##,##0.00") & "    " & "Tasa Libras: " & Format(DtaTasas.Recordset("MontoLibras"), "##,##0.00")
End If

If Not Encontrado Then
  MsgBox "La tasa de esta Fecha no ha sido Grabada"
  Cancel = 100
  Tasa = False
  frmTasa2.Show 1
End If

FrmAuxiliarMovimientos.Show 1
Unload Me
Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

With Me.DtaConsulta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaTasas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

Me.DTPicker1.Value = Format(Now, "dd/mm/yyyy")
Me.CmbMoneda.Text = FrmConciliacion.CmbTipoMoneda.Text
End Sub

