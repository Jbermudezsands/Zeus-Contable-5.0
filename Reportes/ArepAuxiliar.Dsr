VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepAuxiliar 
   Caption         =   "Reporte de Analitico"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepAuxiliar.dsx":0000
End
Attribute VB_Name = "ArepAuxiliar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TransaccionNo As String, FechaTransaccion As Date, Saldo As Double, TotalDebito As Double, TotalCredito As Double

Private Sub ActiveReport_Activate()
QuienReporte = Me.Name
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)
    If Not EOF Then
    'Gets the current records SupplierID
      TransaccionNo = Me.DataControl1.Recordset.Fields("NumeroMovimiento")
      FechaTransaccion = Me.DataControl1.Recordset.Fields("FechaTransaccion")
    End If
End Sub

Private Sub ActiveReport_ReportEnd()
 On Error GoTo err:
   Dim RutaArchivo As String


If FrmReportes.ChkExportar.Value = 1 Then
'    Establecer CancelError a True
'    FrmReportes.CDRuta.CancelError = True
'     Establecer los indicadores
    FrmReportes.CDRuta.Flags = cdlOFNHideReadOnly
'     Establecer los filtros
    FrmReportes.CDRuta.Filter = "Excel (*.XLS)|*.xls"
'     Especificar el filtro predeterminado
'    FrmReportes.CDRuta.FilterIndex = 2
    ' Presentar el cuadro de diálogo Abrir
'    FrmReportes.CDRuta.ShowSave
'    ' Presentar el nombre del archivo seleccionado
'    RutaArchivo = FrmReportes.CDRuta.FileName  'varible que le doy la ruta

   RutaArchivo = FrmReportes.CommonDialog1.FileName + ".xls"
   
    MousePointer = 11
    
    Dim myExportObject As ActiveReportsExcelExport.ARExportExcel
    Dim Nombre As String
    
'    Nombre = InputBox("Digite el Nombre del Archivo", "Sistema de Nominas")
    Set myExportObject = CreateObject("ActiveReportsExcelExport.ARExportExcel")
    myExportObject.FileName = RutaArchivo
    myExportObject.Export Me.Pages
    Set myExportObject = Nothing

End If
err:
    If err.Number <> 0 Then Exit Sub

End Sub

Private Sub ActiveReport_ReportStart()
    On Error GoTo err
    
    QuienReporte = Me.Name
    
    Me.LblMoneda.Caption = Moneda
    
             Me.LblEmpresa = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
             Me.LblEmpresa1 = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
             Me.LblEmpresa2 = "RUC: " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
             Me.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
             Me.LblCodigo.Caption = "Mis Code:" & FrmReportes.DBCodigo.Text
             Me.LblRango.Caption = "Filtrado Desde: " & FrmReportes.CodDesde & " Hasta " & FrmReportes.CodHasta
             If Dir(RutaLogo) <> "" Then
                   Me.Logo.Picture = LoadPicture(RutaLogo)
             End If

    
    Me.Field19.Hyperlink = ""
    
If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    End If
err:
    If err.Number <> 0 Then MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation
    
End Sub

Private Sub Detail_Format()
Dim Mov1 As Double, Mov2 As Double



 If Me.Field26.Text = "0.00" Or Me.Field26.Text = "" Then
   Mov1 = 0
 Else
   Mov1 = Me.Field26.Text
 End If
 
 TotalDebito = Mov1 + TotalDebito
 
  If Me.Field27.Text = "0.00" Or Me.Field27.Text = "" Then
   Mov2 = 0
 Else
   Mov2 = Me.Field27.Text
 End If
 
 TotalCredito = Mov2 + TotalCredito
 
      TipoCuenta = Me.FldTipoCuenta.Text
    

If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
 Saldo = Mov1 - Mov2 + Saldo
    Me.FldSaldo.Text = Format(SaldoIni + Saldo, "###,###,###,##0.#0")
Else
 Saldo = Mov2 - Mov1 + Saldo
    Me.FldSaldo.Text = Format(SaldoIni + Saldo, "###,###,###,##0.#0")
'    Me.LblFinal.Caption = Format(SaldoIni - SaldoFin, "##,##0.00")
End If

'pone al tipo sólo D o C
Me.Field22.Text = Mid(Me.Field22, 1, 1)
If FrmReportes.ChkExportar.Value = 0 Then
  Me.Field19.Hyperlink = Format(FechaTransaccion, "dd/mm/yyyy") & ";" & TransaccionNo
End If
End Sub

Private Sub GroupFooter1_Format()
Dim CodigoCuenta As String, TipoCuenta As String



'///////////////Busco el Acumulado de la cuenta hasta la ultima fecha Seleccionada////////////
'///////////////////////////////////////////////////////////////////////////////////////////////
    NumFecha1 = FrmReportes.DTFecha1.Value
    NumFecha2 = FrmReportes.DTFecha2.Value
    
'    If FrmReportes.DBCodigo.Text = "" Then
'     CodigoCuenta = Me.Field16.Text
'    Else
'     CodigoCuenta = FrmReportes.DBCodigo.Text
'    End If
    
    CodigoCuenta = Me.Field16.Text
    TipoCuenta = Me.FldTipoCuenta.Text
    

'        FrmReportes.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito From Transacciones Where (((Transacciones.FechaTransaccion) <= '" & Format(FrmReportes.DTFecha2.Value, "yyyymmdd") & "')) GROUP BY Transacciones.CodCuentas Having (((Transacciones.CodCuentas) = '" & CodigoCuenta & "'))"
        
      If Moneda = "Cordobas" Then
        FrmReportes.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END) AS MDebito, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END) AS MCredito FROM  Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND  Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo INNER JOIN  Tasas ON IndiceTransaccion.FechaTransaccion = Tasas.FechaTasas " & _
                                                "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING  (Transacciones.CodCuentas = '" & CodigoCuenta & "')"
      Else
        FrmReportes.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN ROUND(Transacciones.Debito,2) ELSE ROUND(Transacciones.Debito / Tasas.MontoCordobas,2) END) AS MDebito, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN ROUND(Transacciones.Credito,2) ELSE ROUND(Transacciones.Credito / Tasas.MontoCordobas,2) END) AS MCredito FROM  Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND  Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo INNER JOIN  Tasas ON IndiceTransaccion.FechaTransaccion = Tasas.FechaTasas " & _
                                                "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING  (Transacciones.CodCuentas = '" & CodigoCuenta & "')"
      End If
        
        FrmReportes.DtaHistorial.Refresh
         
         If Not FrmReportes.DtaHistorial.Recordset.EOF Then
          If Not IsNull(FrmReportes.DtaHistorial.Recordset("MDebito")) Then
           Debito = FrmReportes.DtaHistorial.Recordset("MDebito")
          End If
           If Not IsNull(FrmReportes.DtaHistorial.Recordset("MCredito")) Then
             Credito = FrmReportes.DtaHistorial.Recordset("MCredito")
          End If
          
          If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
            Total = Debito - Credito
            SaldoFin = Total
          Else
            Total = Credito - Debito
            SaldoFin = Total
          End If
                
           
         Else
           SaldoFin = 0
         End If
'
' SaldoIni = Me.LblIni.Caption
'
' SaldoFin = SaldoIni + SaldoFin
' Me.LblFinal = Format(SaldoFin, "##,##0.00")
 Me.LblTotalDebitto.Caption = Format(TotalDebitoAux, "##,##0.00")
 Me.LblTotalCredito.Caption = Format(TotalCreditoAux, "##,##0.00")
 Me.LblFinal = Format(SaldoFinalAuxiliar, "##,##0.00")
End Sub

Private Sub GroupHeader1_Format()
On Error GoTo TipoErrs

Dim CodigoCuenta As String, FechaIni As String, FechaFin As String

TotalDebito = 0
TotalCredito = 0

'      Me.Field16.Visible = True
'      Me.Field17.Visible = True
'      Me.FldTipoCuenta.Visible = True
'      Me.Label20.Visible = True
'      Me.LblIni.Visible = True
'      Me.Label21.Visible = True
'      Me.LblFinal.Visible = True
'      Me.SubReportAuxiliar.Visible = True
      
'///////////////Busco el Acumulado de la cuenta hasta la ultima fecha Seleccionada////////////
'///////////////////////////////////////////////////////////////////////////////////////////////
    SaldoFin = 0
    Debito = 0
    Credito = 0
    SaldoFinalAuxiliar = 0
    TotalDebitoAux = 0
    TotalCreditoAux = 0
    
    NumFecha1 = FrmReportes.DTFecha1.Value
    NumFecha2 = FrmReportes.DTFecha2.Value
    
    FechaIni = Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd")
    FechaFin = Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd")
    
    
    '////////////////////////////////////////////////////////////////////////////
    '//////////////BUSCO EL SALDO INICIAL/////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////
    
    CodigoCuenta = Me.Field16.Text
    TipoCuenta = Me.FldTipoCuenta.Text
    
    If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
      If Moneda = "Cordobas" Then
        FrmReportes.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END) AS MDebito, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END) AS MCredito,IndiceTransaccion.TipoMoneda FROM  Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND  Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo INNER JOIN  Tasas ON IndiceTransaccion.FechaTransaccion = Tasas.FechaTasas " & _
                                                "WHERE (Transacciones.FechaTransaccion < CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas, IndiceTransaccion.TipoMoneda HAVING  (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (IndiceTransaccion.TipoMoneda <> 'Dólares')"
      Else

'        FrmReportes.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END) AS MDebito, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END) AS MCredito,IndiceTransaccion.TipoMoneda FROM  Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND  Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo INNER JOIN  Tasas ON IndiceTransaccion.FechaTransaccion = Tasas.FechaTasas " & _
'                                                "WHERE (Transacciones.FechaTransaccion < CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas, IndiceTransaccion.TipoMoneda HAVING  (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (IndiceTransaccion.TipoMoneda <> 'Córdobas')"
         FrmReportes.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Tasas.MontoCordobas END) AS MDebito, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito END) AS MCredito FROM Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento AND  Transacciones.NPeriodo = IndiceTransaccion.Nperiodo INNER JOIN Tasas ON IndiceTransaccion.FechaTransaccion = Tasas.FechaTasas " & _
                                                 "WHERE (Transacciones.FechaTransaccion < CONVERT(DATETIME,'" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (MAX(IndiceTransaccion.Ajuste) <> N'Córdobas') "
      End If
      
        FrmReportes.DtaHistorial.Refresh
         
         If Not FrmReportes.DtaHistorial.Recordset.EOF Then
          If Not IsNull(FrmReportes.DtaHistorial.Recordset("MDebito")) Then
             Debito = FrmReportes.DtaHistorial.Recordset("MDebito")
          End If
           If Not IsNull(FrmReportes.DtaHistorial.Recordset("MCredito")) Then
             Credito = FrmReportes.DtaHistorial.Recordset("MCredito")
          End If
          Total = Debito - Credito
          SaldoIni = Total
                
           
         Else
           SaldoIni = 0
         End If
     
     Else
     
     If Moneda = "Cordobas" Then
         FrmReportes.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END) AS MDebito, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END) AS MCredito, IndiceTransaccion.TipoMoneda  FROM  Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND  Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo INNER JOIN  Tasas ON IndiceTransaccion.FechaTransaccion = Tasas.FechaTasas " & _
                                                 "WHERE (Transacciones.FechaTransaccion < CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas, IndiceTransaccion.TipoMoneda  HAVING  (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (IndiceTransaccion.TipoMoneda <> 'Dólares')"
'        FrmReportes.DtaHistorial.RecordSource = "SELECT CodCuentas, SUM(Debito * TCambio) AS MDebito, SUM(TCambio * Credito) AS MCredito From Transacciones WHERE     (FechaTransaccion < CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "', 102)) GROUP BY CodCuentas HAVING   (CodCuentas = '" & CodigoCuenta & "')"
        FrmReportes.DtaHistorial.Refresh
     Else
'        FrmReportes.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN ROUND(Transacciones.Debito,2) ELSE ROUND(Transacciones.Debito / Tasas.MontoCordobas,2) END) AS MDebito, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN ROUND(Transacciones.Credito,2) ELSE ROUND(Transacciones.Credito / Tasas.MontoCordobas,2) END) AS MCredito FROM  Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND  Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo INNER JOIN  Tasas ON IndiceTransaccion.FechaTransaccion = Tasas.FechaTasas " & _
'                                                "WHERE (Transacciones.FechaTransaccion < CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING  (Transacciones.CodCuentas = '" & CodigoCuenta & "')"
        FrmReportes.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.TCambio) AS TCambio, SUM(ROUND(Transacciones.Debito, 2)) AS Debito, SUM(ROUND(Transacciones.Credito, 2)) AS Credito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 2)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas, 2)) AS MCredito, SUM(ROUND(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas), 2) - ROUND(Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas), 2)) AS Saldo, IndiceTransaccion.TipoMoneda   FROM Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
                                                "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas, IndiceTransaccion.TipoMoneda  HAVING (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (IndiceTransaccion.TipoMoneda <> 'Córdobas')"
        FrmReportes.DtaHistorial.Refresh
        
     End If
         
         If Not FrmReportes.DtaHistorial.Recordset.EOF Then
          If Not IsNull(FrmReportes.DtaHistorial.Recordset("MDebito")) Then
             Debito = FrmReportes.DtaHistorial.Recordset("MDebito")
          End If
           If Not IsNull(FrmReportes.DtaHistorial.Recordset("MCredito")) Then
             Credito = FrmReportes.DtaHistorial.Recordset("MCredito")
          End If
          Total = Credito - Debito
          SaldoIni = Total
                
           
         Else
           SaldoIni = 0
         End If
     
     
     End If
         
         
         
     Me.LblIni.Caption = Format(SaldoIni, "##,##0.00")
     
  DoEvents
  
 '//////////////////////////////////////////////////////////////////////////////////////////
 '///////////////////////ASIGNO EL SUBREPORTE/////////////////////////////////////////////////////
 '////////////////////////////////////////////////////////////////////////////////////////

 If Moneda = "Cordobas" Then
 
 Sql = "SELECT  Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.TCambio, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTransaccion, Cuentas.TipoCuenta, Transacciones.TCambio AS Expr1, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END AS Debito, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END AS Credito, IndiceTransaccion.Nperiodo AS Expr2, Tasas.MontoCordobas FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND " & _
       "Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo INNER JOIN Tasas ON IndiceTransaccion.FechaTransaccion = Tasas.FechaTasas " & _
       "WHERE (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (IndiceTransaccion.Ajuste <> 'Dólares') AND (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & FechaIni & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"

'    SQL = "SELECT  Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.NumeroMovimiento,Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.TCambio, Transacciones.FacturaNo,Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTransaccion, Cuentas.TipoCuenta, Transacciones.TCambio AS Expr1,Transacciones.TCambio * Transacciones.Debito AS Debito, Transacciones.TCambio * Transacciones.Credito AS Credito  " & _
'          "FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas " & _
'          "WHERE (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & FechaIni & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"
 
 Else
'    Sql = "SELECT  Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.TCambio, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTransaccion, Cuentas.TipoCuenta, Transacciones.TCambio AS Expr1, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN ROUND(Transacciones.Debito,2) ELSE ROUND(Transacciones.Debito / Tasas.MontoCordobas,2) END AS Debito, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN ROUND(Transacciones.Credito,2) ELSE ROUND(Transacciones.Credito / Tasas.MontoCordobas,2) END AS Credito, IndiceTransaccion.Nperiodo AS Expr2, Tasas.MontoCordobas FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND " & _
'           "Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo INNER JOIN Tasas ON IndiceTransaccion.FechaTransaccion = Tasas.FechaTasas " & _
'           "WHERE (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (IndiceTransaccion.TipoMoneda <> 'Córdobas') AND (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & FechaIni & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"
     
    Sql = "SELECT  Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.TCambio, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTransaccion, Cuentas.TipoCuenta, Transacciones.TCambio AS Expr1, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Tasas.MontoCordobas END AS Debito, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Tasas.MontoCordobas END AS Credito, IndiceTransaccion.Nperiodo AS Expr2, Tasas.MontoCordobas FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND " & _
          "Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo INNER JOIN Tasas ON IndiceTransaccion.FechaTransaccion = Tasas.FechaTasas " & _
          "WHERE (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (IndiceTransaccion.Ajuste <> 'Córdobas') AND (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & FechaIni & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"
     
 End If
   
       Set Me.SubReportAuxiliar.object = New ArepAuxiliarSrpt
       Me.SubReportAuxiliar.object.DataControl1.ConnectionString = ConexionReporte
       Me.SubReportAuxiliar.object.DataControl1.Source = Sql
      
       
'    Else
''      Me.Field16.Visible = False
''      Me.Field17.Visible = False
''      Me.FldTipoCuenta.Visible = False
''      Me.Label20.Visible = False
''      Me.LblIni.Visible = False
''      Me.Label21.Visible = False
''      Me.LblFinal.Visible = False
''      Me.SubReportAuxiliar.Visible = False
''
'    End If
    
    DoEvents
    
    
Exit Sub
TipoErrs:
 MsgBox err.Description
     
End Sub

