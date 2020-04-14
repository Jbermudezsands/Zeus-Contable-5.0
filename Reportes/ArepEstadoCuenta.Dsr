VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepEstadoCuenta 
   Caption         =   "Estado de cuenta"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20340
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35878
   _ExtentY        =   19420
   SectionData     =   "ArepEstadoCuenta.dsx":0000
End
Attribute VB_Name = "ArepEstadoCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FacturaNo As String, FechaFact As Date
Private Sub ActiveReport_FetchData(EOF As Boolean)
    If Not EOF Then
    'Gets the current records SupplierID
      FacturaNo = Me.DataControl1.Recordset.Fields("FacturaNo")
'      FechaFact = Me.DataControl1.Recordset.Fields("FacturaNo")
    End If
End Sub

Private Sub ActiveReport_hyperLink(ByVal Button As Integer, Link As String)
 Dim FacturaNumero As String, FechaFactura As Date, SQL As String
 Dim rpt As Object, fPreview As New FrmPreview
'Check to see if an email link or web page has been selected

FacturaNumero = Link
If InStr(1, Link, "htm", vbTextCompare) = 0 And InStr(1, Link, "mailto", vbTextCompare) = 0 Then

         If MDIPrimero.AdoConfiguracion.Recordset.RecordCount > 0 Then
             If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!ConexionFacturacion) Then
                ConexionFacturacion = MDIPrimero.AdoConfiguracion.Recordset!ConexionFacturacion
             Else
                ConexionFacturacion = ""
             End If
         End If


    FacturaNo = Link
    
    FechaFactura = Me.Field6.Text
    
    '---------------------------------------------------------------------------------------------------
    '-----------------------------------------CONSULTO LOS DATOS DE LA FACTURA ------------------------
    '---------------------------------------------------------------------------------------------------
    SQL = "SELECT  * FROM  Facturas INNER JOIN Bodegas ON Facturas.Cod_Bodega = Bodegas.Cod_Bodega " & _
          "WHERE (Facturas.Numero_Factura = '" & FacturaNumero & "') AND (Facturas.Fecha_Factura = CONVERT(DATETIME, '" & Format(FechaFactura, "yyyy-mm-dd") & "', 102)) AND (Facturas.Tipo_Factura = N'Factura')"
  
    MDIPrimero.AdoConsultaFacturacion.ConnectionString = ConexionFacturacion
    MDIPrimero.AdoConsultaFacturacion.RecordSource = SQL
    MDIPrimero.AdoConsultaFacturacion.Refresh
    If Not MDIPrimero.AdoConsultaFacturacion.Recordset.EOF Then
      ArepFacturas.LblSubTotal.Caption = Format(MDIPrimero.AdoConsultaFacturacion.Recordset("SubTotal"), "##,##0.00")
      ArepFacturas.LblIVA.Caption = Format(MDIPrimero.AdoConsultaFacturacion.Recordset("IVA"), "##,##0.00")
      ArepFacturas.LblTotal.Caption = Format(MDIPrimero.AdoConsultaFacturacion.Recordset("SubTotal") + MDIPrimero.AdoConsultaFacturacion.Recordset("IVA"), "##,##0.00")
      ArepFacturas.LblBodega.Caption = MDIPrimero.AdoConsultaFacturacion.Recordset("Cod_Bodega") & " " & MDIPrimero.AdoConsultaFacturacion.Recordset("Nombre_Bodega")
      ArepFacturas.LblObservaciones.Caption = MDIPrimero.AdoConsultaFacturacion.Recordset("Observaciones")
    End If
    

    
    
    SQL = "SELECT  Productos.*, Detalle_Facturas.* FROM Detalle_Facturas INNER JOIN Productos ON Detalle_Facturas.Cod_Producto = Productos.Cod_Productos  " & _
          "WHERE  (Detalle_Facturas.Numero_Factura = '" & FacturaNumero & "') AND (Detalle_Facturas.Tipo_Factura = 'Factura') AND (Detalle_Facturas.Fecha_Factura = CONVERT(DATETIME, '" & Format(FechaFactura, "yyyy-mm-dd") & "', 102))"
    
     
      ArepFacturas.DataControl1.ConnectionString = ConexionFacturacion
      ArepFacturas.DataControl1.Source = SQL
    
       ArepFacturas.Logo.Picture = LoadPicture(RutaLogo)
    
      ArepFacturas.LblEmpresa.Caption = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
      ArepFacturas.LblEmpresa1.Caption = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
      ArepFacturas.LblEmpresa2.Caption = "RUC " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
      ArepFacturas.LblCodigoCliente.Caption = Me.Field1.Text
      ArepFacturas.LblNombreCliente.Caption = Me.Field2.Text
      
'      ArepFacturas.Show 1
      
   
         Set rpt = New ArepFacturas
         rpt.DataControl1.ConnectionString = ConexionFacturacion
         rpt.DataControl1.Source = SQL
         fPreview.RunReport rpt
         fPreview.Show 1


End If
End Sub

Private Sub ActiveReport_ReportStart()
  QuienReporte = Me.Name
  
      Me.Logo.Picture = LoadPicture(RutaLogo)
      Me.LblEmpresa.Caption = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
      Me.LblEmpresa1.Caption = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
      Me.LblEmpresa2.Caption = "RUC " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
End Sub

Private Sub Detail_Format()
Dim FechaVence As String

FechaVence = Format(Me.Field7.Text, "dd/mm/yyyy")
If FechaVence = "01/01/1900" Then
  Me.Field7.Text = Me.Field6.Text
End If



Me.Field5.Hyperlink = FacturaNo
End Sub

Private Sub PageHeader_Format()
 Dim CodCliente As String, Sqlconsulta As String
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
     Sqlconsulta = "SELECT Transacciones.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS Debito, SUM(Transacciones.Credito * Transacciones.TCambio) AS Credito, SUM((Transacciones.Debito - Transacciones.Credito) * Transacciones.TCambio) AS Saldo, SUM(Transacciones.TCambio) AS Expr5 FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas " & _
                   "WHERE (Transacciones.FechaTransaccion < CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas, Cuentas.TipoCuenta HAVING (Transacciones.CodCuentas = '" & CodCliente & "') AND (Cuentas.TipoCuenta = N'Cuentas x Cobrar')"
     FrmReportes.DtaConsulta.RecordSource = Sqlconsulta
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
     Sqlconsulta = "SELECT Transacciones.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS Debito, SUM(Transacciones.Credito * Transacciones.TCambio) AS Credito, SUM((Transacciones.Debito - Transacciones.Credito) * Transacciones.TCambio) AS Saldo, SUM(Transacciones.TCambio) AS Expr5 FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas " & _
                   "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas, Cuentas.TipoCuenta HAVING (Transacciones.CodCuentas = '" & CodCliente & "') AND (Cuentas.TipoCuenta = N'Cuentas x Cobrar')"
    
     FrmReportes.DtaConsulta.RecordSource = Sqlconsulta
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
''    Fecha2 = Format(FechaFin, "yyyy-mm-dd")
'    SQL = "SELECT     Transacciones.CodCuentas AS CodCuentas, Transacciones.DescripcionMovimiento AS DescripcionMovimiento, " & _
'         "Transacciones.FechaTransaccion AS FechaTransaccion, Transacciones.FechaVence AS FechaVence, " & _
'         "Transacciones.NumeroMovimiento AS NumeroMovimiento, Cuentas.TipoCuenta AS TipoCuenta, Transacciones.FacturaNo AS FacturaNo, " & _
'         "Transacciones.Debito * Transacciones.TCambio AS Debito, Transacciones.Credito * Transacciones.TCambio AS Credito, " & _
'         "(Transacciones.Debito - Transacciones.Credito) * Transacciones.TCambio AS Saldo, Transacciones.ChequeNo, Cuentas.DescripcionCuentas, " & _
'         "Transacciones.TCambio " & _
'         "FROM         Transacciones INNER JOIN " & _
'         "Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas " & _
'         "WHERE     (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND ((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "') " & _
'         "ORDER BY Transacciones.FacturaNo,Transacciones.FechaTransaccion "
'
'    Set Me.SubReport.object = New EstadoCuentaSrpt
'
'    Me.SubReport.object.AdoEstadoCuenta.ConnectionString = ConexionReporte
'    Me.SubReport.object.AdoEstadoCuenta.Source = SQL
'
    
    
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
