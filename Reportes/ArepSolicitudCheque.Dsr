VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepSolicitudCheque 
   Caption         =   "Solicitud de Cheque"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepSolicitudCheque.dsx":0000
End
Attribute VB_Name = "ArepSolicitudCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public TipoMoneda As String, NumeroSolicitud As String, Periodo As Double, FechaSolicitud

Private Sub ActiveReport_ReportStart()
    On Error GoTo err
If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    End If
err:
    If err.Number <> 0 Then MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation
    
    
   Me.LblEmpresa = MDIPrimero.AdoConfiguracion.Recordset("NombreEmpresa")
   Me.LblEmpresa1 = MDIPrimero.AdoConfiguracion.Recordset("Direccion")
   Me.LblEmpresa2 = "RUC: " & MDIPrimero.AdoConfiguracion.Recordset("NumeroRuc")
   
    Set ew = New cls_NumEnglishWord
    Set sw = New cls_NumSpanishWord
End Sub

Private Sub GroupFooter1_Format()
 Dim PresupuestoAnual As Double, KeyPresupuesto As String, MontoReal As Double, MontoSolicitud As Double
 
 MontoSolicitud = 0
 If Me.FldMonto.Text <> "" Then
  MontoSolicitud = Me.FldMonto.Text
 End If
 
 MDIPrimero.AdoConsulta.RecordSource = "SELECT  TransaccionesSolicitudPago.CodCuentas, TransaccionesSolicitudPago.NumeroMovimiento, TransaccionesSolicitudPago.TCambio, TransaccionesSolicitudPago.Debito, TransaccionesSolicitudPago.Credito, TransaccionesSolicitudPago.KeyPresupuesto, TransaccionesSolicitudPago.Presupuesto, EstructuraPresupuesto.DescripcionGrupo , PresupuestoAnual.MontoAnual, TransaccionesSolicitudPago.NPeriodo FROM TransaccionesSolicitudPago INNER JOIN EstructuraPresupuesto ON TransaccionesSolicitudPago.KeyPresupuesto = EstructuraPresupuesto.KeyGrupo INNER JOIN PresupuestoAnual ON EstructuraPresupuesto.KeyGrupo = PresupuestoAnual.CodigoCuenta   " & _
                                       "Where (TransaccionesSolicitudPago.NumeroMovimiento = '" & NumeroSolicitud & "') And (TransaccionesSolicitudPago.NPeriodo = " & Periodo & " )"
 MDIPrimero.AdoConsulta.Refresh
 If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
   Me.LblDescripcion.Caption = MDIPrimero.AdoConsulta.Recordset("DescripcionGrupo")
   If Not IsNull(MDIPrimero.AdoConsulta.Recordset("KeyPresupuesto")) Then
      KeyPresupuesto = MDIPrimero.AdoConsulta.Recordset("KeyPresupuesto")
      
      PresupuestoAnual = 0
      MontoReal = 0
      '////////////////////////////////////BUSCO EL MONTO DEL PRESUPUESTO ANUAL /////////////////
      MDIPrimero.AdoConsulta.RecordSource = "SELECT NumeroTabla, CodigoCuenta, MontoAnual From PresupuestoAnual WHERE  (CodigoCuenta = '" & KeyPresupuesto & "')"
      MDIPrimero.AdoConsulta.Refresh
      If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
        PresupuestoAnual = MDIPrimero.AdoConsulta.Recordset("MontoAnual")
      End If
      
      
      MDIPrimero.AdoConsulta.RecordSource = "SELECT Transacciones.KeyPresupuesto, SUM(Transacciones.TCambio * Transacciones.Debito - Transacciones.TCambio * Transacciones.Credito) AS Saldo, Transacciones.Presupuesto FROM Transacciones INNER JOIN IndiceSolicitudPago ON Transacciones.NumeroMovimiento = IndiceSolicitudPago.NumeroTransaccion AND Transacciones.NPeriodo = IndiceSolicitudPago.NPeriodoTransaccion WHERE  (Transacciones.KeyPresupuesto = '" & KeyPresupuesto & "') AND (IndiceSolicitudPago.NumeroMovimiento <> " & NumeroSolicitud & ") AND (Transacciones.FechaTransaccion <= CONVERT(DATETIME,'" & Format(FechaSolicitud, "yyyy-MM-dd") & "', 102)) GROUP BY Transacciones.Presupuesto, Transacciones.KeyPresupuesto"
      MDIPrimero.AdoConsulta.Refresh
      If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
        MontoReal = MDIPrimero.AdoConsulta.Recordset("Saldo")
      End If
      
      Me.LblSaldoInicial.Caption = Format(PresupuestoAnual - MontoReal, "##,##0.00")
      Me.LblMontoSolicitud.Caption = Format(MontoSolicitud, "##,##0.00")
      Me.LblSaldoFinal.Caption = Format(PresupuestoAnual - MontoReal - MontoSolicitud, "##,##0.00")
      
      
      
      
      
   End If
   
   
 End If
 
 
 
End Sub

Private Sub PageHeader_Format()
 Dim Monto As Double, Letras As String
 
 Monto = 0
 If Me.FldMonto.Text <> "" Then
  Monto = Me.FldMonto.Text
 End If
 
'
'            If TipoMoneda = "Dólares" Then
'             Letras = sw.ConvertCurrencyToSpanish(Monto, "Dólares")
'            ElseIf TipoMoneda = "Córdobas" Then
'             Letras = sw.ConvertCurrencyToSpanish(Monto, "Córdobas")
'            End If
            
            
 Me.LblDescripcionMonto.Caption = Letras
 
End Sub
