VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} EstadoCuenta 
   Caption         =   "Reporte de Estdo de Cuetas"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20340
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   35878
   _ExtentY        =   19420
   SectionData     =   "EstadoCuenta.dsx":0000
End
Attribute VB_Name = "EstadoCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ActiveReport_ReportEnd()
 On Error GoTo err:
   Dim RutaArchivo As String
    Dim myExportObject As ActiveReportsExcelExport.ARExportExcel
    Set myExportObject = CreateObject("ActiveReportsExcelExport.ARExportExcel")
If FrmReportes.ChkExportar.Value = 1 Then
   ' Establecer CancelError a True
    FrmReportes.CDRuta.CancelError = True
    ' Establecer los indicadores
    FrmReportes.CDRuta.Flags = cdlOFNHideReadOnly
    ' Establecer los filtros
    FrmReportes.CDRuta.Filter = "Excel (*.XLS)|*.xls"
    ' Especificar el filtro predeterminado
    FrmReportes.CDRuta.FilterIndex = 2
    ' Presentar el cuadro de diálogo Abrir
    FrmReportes.CDRuta.ShowSave
    ' Presentar el nombre del archivo seleccionado
    RutaArchivo = FrmReportes.CDRuta.FileName 'varible que le doy la ruta
   
    MousePointer = 11
    myExportObject.FileName = RutaArchivo
    myExportObject.Export Me.Pages
    Set myExportObject = Nothing
    MousePointer = 1
End If
err:
    If err.Number <> 0 Then Exit Sub

End Sub
Private Sub ActiveReport_ReportStart()
    On Error GoTo err
If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    End If
err:
    If err.Number <> 0 Then MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation
    


  Me.LblFecha1.Caption = FrmReportes.DTFecha1.Value
  Me.LblFecha.Caption = FrmReportes.DTFecha2.Value
  
  Me.LblEmpresa.Caption = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
  Me.LblEmpresa1.Caption = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
  Me.LblEmpresa2.Caption = "RUC " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")

End Sub

Private Sub GroupHeader1_Format()
 Dim CodCliente As String, sqlconsulta As String
 Dim FechaIni As String, SaldoCuenta As Double, FechaFin As String, MontoTotal As Double
 Dim FechaVence As Date, Cantdias As Double
 Dim SaldoVencer As Double, SaldoMayor30 As Double, SaldoMayor60 As Double
 

     FechaIni = Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd")
     FechaFin = Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd")

     
  '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  '//////////////////////BUSCO EL SALDO INICIAL////////////////////////////////////////////////////////////////////////////
  '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
     
     CodCliente = Me.Field1.Text
'     Sqlconsulta = "SELECT  Transacciones.CodCuentas AS CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.FechaVence)AS FechaVence, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Transacciones.FacturaNo) AS FacturaNo, SUM(Transacciones.Debito * Transacciones.TCambio) AS Debito, SUM(Transacciones.Credito * Transacciones.TCambio) AS Credito, SUM((Transacciones.Debito - Transacciones.Credito) * Transacciones.TCambio) AS Saldo, AVG(Transacciones.TCambio) AS TCambio " & _
'                   "FROM  Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas GROUP BY Transacciones.CodCuentas " & _
'                   "HAVING (MAX(Transacciones.FechaTransaccion) < CONVERT(DATETIME, '" & FechaIni & "', 102)) AND (MAX(Cuentas.TipoCuenta) = 'Cuentas x Cobrar') AND (MAX(Transacciones.FacturaNo) <> N'.') AND (MAX(Transacciones.FacturaNo) <> N'-') AND (Transacciones.CodCuentas = '" & CodCliente & "') ORDER BY MAX(Transacciones.FacturaNo)"
     sqlconsulta = "SELECT Transacciones.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS Debito, SUM(Transacciones.Credito * Transacciones.TCambio) AS Credito, SUM((Transacciones.Debito - Transacciones.Credito) * Transacciones.TCambio) AS Saldo, SUM(Transacciones.TCambio) AS Expr5 FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas " & _
                   "WHERE (Transacciones.FechaTransaccion < CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas, Cuentas.TipoCuenta HAVING (Transacciones.CodCuentas = '" & CodCliente & "') AND (Cuentas.TipoCuenta = N'Cuentas x Cobrar')"
     FrmReportes.DtaConsulta.RecordSource = sqlconsulta
     FrmReportes.DtaConsulta.Refresh
     If Not FrmReportes.DtaConsulta.Recordset.EOF Then
       SaldoCuenta = FrmReportes.DtaConsulta.Recordset("Saldo")
     Else
       SaldoCuenta = 0
     End If
    
     Me.LblSaldo.Caption = Format(SaldoCuenta, "##,##0.00")
     
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '/////////////// BUSCO EL SALDO FINAL//////////////////////////////////////////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
      CodCliente = Me.Field1.Text
'     Sqlconsulta = "SELECT  Transacciones.CodCuentas AS CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.FechaVence)AS FechaVence, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Transacciones.FacturaNo) AS FacturaNo, SUM(Transacciones.Debito * Transacciones.TCambio) AS Debito, SUM(Transacciones.Credito * Transacciones.TCambio) AS Credito, SUM((Transacciones.Debito - Transacciones.Credito) * Transacciones.TCambio) AS Saldo, AVG(Transacciones.TCambio) AS TCambio " & _
'                   "FROM  Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas GROUP BY Transacciones.CodCuentas " & _
'                   "HAVING (MAX(Transacciones.FechaTransaccion) < CONVERT(DATETIME, '" & FechaFin & "', 102)) AND (MAX(Cuentas.TipoCuenta) = 'Cuentas x Cobrar') AND (MAX(Transacciones.FacturaNo) <> N'.') AND (MAX(Transacciones.FacturaNo) <> N'-') AND (Transacciones.CodCuentas = '" & CodCliente & "') ORDER BY MAX(Transacciones.FacturaNo)"
     sqlconsulta = "SELECT Transacciones.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS Debito, SUM(Transacciones.Credito * Transacciones.TCambio) AS Credito, SUM((Transacciones.Debito - Transacciones.Credito) * Transacciones.TCambio) AS Saldo, SUM(Transacciones.TCambio) AS Expr5 FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas " & _
                   "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas, Cuentas.TipoCuenta HAVING (Transacciones.CodCuentas = '" & CodCliente & "') AND (Cuentas.TipoCuenta = N'Cuentas x Cobrar')"
    
     FrmReportes.DtaConsulta.RecordSource = sqlconsulta
     FrmReportes.DtaConsulta.Refresh
     If Not FrmReportes.DtaConsulta.Recordset.EOF Then
      MontoTotal = FrmReportes.DtaConsulta.Recordset("Saldo")
     Else
      MontoTotal = 0
     End If
 
     Me.LblMontoTotal.Caption = Format(MontoTotal, "##,##0.00")
 
 '//////////////////////////////////////////////////////////////////////////////////////////
 '///////////////////////ASIGNO EL SUBREPORTE/////////////////////////////////////////////////////
 '////////////////////////////////////////////////////////////////////////////////////////
 
    CodigoCuenta = Me.Field1.Text
'    Fecha2 = Format(FechaFin, "yyyy-mm-dd")
    SQL = "SELECT     Transacciones.CodCuentas AS CodCuentas, Transacciones.DescripcionMovimiento AS DescripcionMovimiento, " & _
         "Transacciones.FechaTransaccion AS FechaTransaccion, Transacciones.FechaVence AS FechaVence, " & _
         "Transacciones.NumeroMovimiento AS NumeroMovimiento, Cuentas.TipoCuenta AS TipoCuenta, Transacciones.FacturaNo AS FacturaNo, " & _
         "Transacciones.Debito * Transacciones.TCambio AS Debito, Transacciones.Credito * Transacciones.TCambio AS Credito, " & _
         "(Transacciones.Debito - Transacciones.Credito) * Transacciones.TCambio AS Saldo, Transacciones.ChequeNo, Cuentas.DescripcionCuentas, " & _
         "Transacciones.TCambio " & _
         "FROM         Transacciones INNER JOIN " & _
         "Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas " & _
         "WHERE     (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND ((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "') " & _
         "ORDER BY Transacciones.FacturaNo,Transacciones.FechaTransaccion "
 
    Set Me.SubReport.object = New EstadoCuentaSrpt
    
    Me.SubReport.object.AdoEstadoCuenta.ConnectionString = ConexionReporte
    Me.SubReport.object.AdoEstadoCuenta.Source = SQL
    
    
    
 '//////////////////////////////////////////////////////////////////////////////////////////////////////////
 '/////////////////////UBICO EL SALDO SEGUN SU VENCIMIENTO//////////////////////////////////////////////////
 '//////////////////////////////////////////////////////////////////////////////////////////////////////////
 
    SQL = "SELECT Transacciones.CodCuentas AS CodCuentas, MIN(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.FechaVence) AS FechaVence, Transacciones.FacturaNo AS FacturaNo, SUM(Transacciones.Debito * Transacciones.TCambio) AS Debito, SUM(Transacciones.Credito * Transacciones.TCambio) AS Credito, SUM((Transacciones.Debito - Transacciones.Credito) * Transacciones.TCambio) AS Saldo  " & _
          "FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas  GROUP BY Transacciones.CodCuentas, Transacciones.FacturaNo  " & _
          "HAVING (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (MIN(Transacciones.FechaTransaccion) <= CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY Transacciones.FacturaNo, MIN(Transacciones.FechaTransaccion)"
     FrmReportes.DtaConsulta.RecordSource = SQL
     FrmReportes.DtaConsulta.Refresh
     
     SaldoVencer = 0
     Saldo30 = 0
     SaldoMayor30 = 0
     SaldoMayor60 = 0
     
     
     Do While Not FrmReportes.DtaConsulta.Recordset.EOF
       FechaVence = FrmReportes.DtaConsulta.Recordset("FechaVence")
       Cantdias = CDate(FechaFin) - CDate(FechaVence)
       If Cantdias >= -10 And Cantdias <= 0 Then
        If Not IsNull(FrmReportes.DtaConsulta.Recordset("Saldo")) Then
         SaldoVencer = SaldoVencer + FrmReportes.DtaConsulta.Recordset("Saldo")
        End If
       ElseIf Cantdias >= 1 And Cantdias <= 30 Then
        If Not IsNull(FrmReportes.DtaConsulta.Recordset("Saldo")) Then
         Saldo30 = Saldo30 + FrmReportes.DtaConsulta.Recordset("Saldo")
        End If
       ElseIf Cantdias >= 31 And Cantdias <= 60 Then
        If Not IsNull(FrmReportes.DtaConsulta.Recordset("Saldo")) Then
         SaldoMayor30 = SaldoMayor30 + FrmReportes.DtaConsulta.Recordset("Saldo")
        End If
       ElseIf Cantdias >= 61 Then
        If Not IsNull(FrmReportes.DtaConsulta.Recordset("Saldo")) Then
         SaldoMayor60 = SaldoMayor60 + FrmReportes.DtaConsulta.Recordset("Saldo")
        End If
        
       
       End If
       
       
       FrmReportes.DtaConsulta.Recordset.MoveNext
     Loop
     
  Me.LblVencer.Text = Format(SaldoVencer, "##,##0.00")
  Me.LblMayor1.Text = Format(Saldo30, "##,##0.00")
  Me.LblMayor30.Text = Format(SaldoMayor30, "##,##0.00")
  Me.LblMayor60.Text = Format(SaldoMayor60, "##,##0.00")
End Sub
 
