VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmPreview 
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   360
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arv 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7646
      SectionData     =   "FrmPreview.frx":0000
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu ExportaPDF 
         Caption         =   "&Exporta  PDF"
      End
      Begin VB.Menu ExportaExcel 
         Caption         =   "&Exportar Excel"
      End
   End
End
Attribute VB_Name = "FrmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub arv_hyperLink(ByVal Button As Integer, Link As String)
On Error GoTo TipoErrs
 Dim FacturaNumero As String, FechaFactura As Date, SQL As String, Fuente As String
 Dim FechaTransaccion As Date, NumeroTransaccion As Double
 Dim rpt As Object, fPreview As New FrmPreview
 Dim CodigoCuentaDesde As String, CodigoCuentaHasta As String
'Check to see if an email link or web page has been selected



  Select Case QuienReporte
     Case "ArepBalanza"
        If InStr(1, Link, "htm", vbTextCompare) = 0 And InStr(1, Link, "mailto", vbTextCompare) = 0 Then
              ArepAuxiliar.LblRangoFecha = "Desde " & FrmReportes.DTFecha1.Value & " Hasta " & FrmReportes.DTFecha2.Value
              ArepAuxiliar.LblFecha = Format(Now, "dd/mm/yyyy")
              ArepAuxiliar.DataControl1.ConnectionString = ConexionReporte
          
          
              
                
                CodigoCuentaDesde = LeeCadena(Link, 1)
                CodigoCuentaHasta = CodigoCuentaDesde
                
              SQL = "SELECT Transacciones.CodCuentas,  MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Transacciones.NPeriodo) AS NPeriodo,MAX(Transacciones.NTransaccion) AS NTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.VoucherNo) AS VoucherNo, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.Clave) AS Clave, SUM(Transacciones.Debito) AS Debito, SUM(Transacciones.Credito) AS Credito, MAX(Transacciones.FacturaNo) AS FacturaNo, MAX(Transacciones.ChequeNo) AS ChequeNo, MAX(Transacciones.Fuente) AS Fuente, MAX(Cuentas.TipoCuenta) AS TipoCuenta, SUM(Transacciones.Debito + Transacciones.Credito) As Saldo FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE  (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas " & _
                                                   "HAVING (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') AND (SUM(Transacciones.Debito + Transacciones.Credito) <> 0) ORDER BY Transacciones.CodCuentas"
                
    '         SQL = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.TCambio, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTransaccion, Cuentas.TipoCuenta, Transacciones.TCambio AS Expr1, Transacciones.TCambio * Transacciones.Debito AS Debito, Transacciones.TCambio * Transacciones.Credito AS Credito FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
    '                     "WHERE (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') AND (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME,'" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"
                   ArepAuxiliar.DataControl1.ConnectionString = ConexionReporte
    
                 
                Set rpt = New ArepAuxiliar
                rpt.DataControl1.ConnectionString = ConexionReporte
                rpt.DataControl1.Source = SQL
                fPreview.RunReport rpt
                fPreview.Show 1
                
                QuienReporte = "ArepBalanza"
        End If
     
     
     
     Case "ArepComprobanteDiario"
           If InStr(1, Link, "htm", vbTextCompare) = 0 And InStr(1, Link, "mailto", vbTextCompare) = 0 Then
                  ArepAuxiliar.LblRangoFecha = "Desde " & FrmReportes.DTFecha1.Value & " Hasta " & FrmReportes.DTFecha2.Value
                  ArepAuxiliar.LblFecha = Format(Now, "dd/mm/yyyy")
                  ArepAuxiliar.DataControl1.ConnectionString = ConexionReporte
              
              
                  
                    
                    CodigoCuentaDesde = LeeCadena(Link, 1)
                    CodigoCuentaHasta = CodigoCuentaDesde
                    
                  
                    
                  SQL = "SELECT Transacciones.CodCuentas,  MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Transacciones.NPeriodo) AS NPeriodo,MAX(Transacciones.NTransaccion) AS NTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.VoucherNo) AS VoucherNo, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.Clave) AS Clave, SUM(Transacciones.Debito) AS Debito, SUM(Transacciones.Credito) AS Credito, MAX(Transacciones.FacturaNo) AS FacturaNo, MAX(Transacciones.ChequeNo) AS ChequeNo, MAX(Transacciones.Fuente) AS Fuente, MAX(Cuentas.TipoCuenta) AS TipoCuenta, SUM(Transacciones.Debito + Transacciones.Credito) As Saldo FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE  (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas " & _
                                                       "HAVING (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') AND (SUM(Transacciones.Debito + Transacciones.Credito) <> 0) ORDER BY Transacciones.CodCuentas"
                    
        '         SQL = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.TCambio, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTransaccion, Cuentas.TipoCuenta, Transacciones.TCambio AS Expr1, Transacciones.TCambio * Transacciones.Debito AS Debito, Transacciones.TCambio * Transacciones.Credito AS Credito FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
        '                     "WHERE (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') AND (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME,'" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"
                       ArepAuxiliar.DataControl1.ConnectionString = ConexionReporte
        
                     
                    Set rpt = New ArepAuxiliar
                    rpt.DataControl1.ConnectionString = ConexionReporte
                    rpt.DataControl1.Source = SQL
                    fPreview.RunReport rpt
                    fPreview.Show 1
                    
                    QuienReporte = "ArepComprobanteDiario"
            End If
     
  
  
  
      Case "ArepTransacciones"
      
      
      

         
            If InStr(1, Link, "htm", vbTextCompare) = 0 And InStr(1, Link, "mailto", vbTextCompare) = 0 Then
            
                     If MDIPrimero.AdoConfiguracion.Recordset.RecordCount > 0 Then
                         If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!ConexionFacturacion) Then
                            ConexionFacturacion = MDIPrimero.AdoConfiguracion.Recordset!ConexionFacturacion
                         Else
                            ConexionFacturacion = ""
                         End If
                     End If
                     
                     If ConexionFacturacion = "" Then
                       Exit Sub
                     End If
                     
                        FechaFactura = LeeCadena(Link, 1)
                        FacturaNumero = LeeCadena(Link, 2)
                        NumeroFact = FacturaNumero
                        Fuente = LeeCadena(Link, 3)
                        If Fuente = "VTAS" Then
                          Tipo = "Factura"
                        ElseIf Fuente = "Salida Bodega" Then
                          Tipo = "Salida Bodega"
                        End If
                           
'                FacturaNo = Link
                
'                FechaFactura = ArepEstadoCuenta.Field6.Text
                
                '---------------------------------------------------------------------------------------------------
                '-----------------------------------------CONSULTO LOS DATOS DE LA FACTURA ------------------------
                '---------------------------------------------------------------------------------------------------
                SQL = "SELECT  * FROM  Facturas INNER JOIN Bodegas ON Facturas.Cod_Bodega = Bodegas.Cod_Bodega " & _
                      "WHERE (Facturas.Numero_Factura = '" & FacturaNumero & "') AND (Facturas.Fecha_Factura = CONVERT(DATETIME, '" & Format(FechaFactura, "yyyy-mm-dd") & "', 102)) AND (Facturas.Tipo_Factura = '" & Tipo & "')"
'                SQL = "SELECT  * FROM  Facturas INNER JOIN Bodegas ON Facturas.Cod_Bodega = Bodegas.Cod_Bodega " & _
'                      "WHERE (Facturas.Numero_Factura = '" & FacturaNumero & "') AND (Facturas.Tipo_Factura = '" & Tipo & "')"
              
                MDIPrimero.AdoConsultaFacturacion.ConnectionString = ConexionFacturacion
                MDIPrimero.AdoConsultaFacturacion.RecordSource = SQL
                MDIPrimero.AdoConsultaFacturacion.Refresh
                If Not MDIPrimero.AdoConsultaFacturacion.Recordset.EOF Then
                  ArepFacturas.LblSubTotal.Caption = Format(MDIPrimero.AdoConsultaFacturacion.Recordset("SubTotal"), "##,##0.00")
                  ArepFacturas.LblIVA.Caption = Format(MDIPrimero.AdoConsultaFacturacion.Recordset("IVA"), "##,##0.00")
                  ArepFacturas.LblTotal.Caption = Format(MDIPrimero.AdoConsultaFacturacion.Recordset("SubTotal") + MDIPrimero.AdoConsultaFacturacion.Recordset("IVA"), "##,##0.00")
                  ArepFacturas.LblBodega.Caption = MDIPrimero.AdoConsultaFacturacion.Recordset("Cod_Bodega") & " " & MDIPrimero.AdoConsultaFacturacion.Recordset("Nombre_Bodega")
                  ArepFacturas.LblObservaciones.Caption = MDIPrimero.AdoConsultaFacturacion.Recordset("Observaciones")
                  ArepFacturas.LblCodigoCliente.Caption = MDIPrimero.AdoConsultaFacturacion.Recordset("Cod_Cliente")
                  ArepFacturas.LblNombreCliente.Caption = MDIPrimero.AdoConsultaFacturacion.Recordset("Nombre_Cliente")
                  ArepFacturas.LblNuestraRef.Caption = MDIPrimero.AdoConsultaFacturacion.Recordset("MonedaFactura")
                End If
                
            
                
                
'                SQL = "SELECT  Productos.*, Detalle_Facturas.* FROM Detalle_Facturas INNER JOIN Productos ON Detalle_Facturas.Cod_Producto = Productos.Cod_Productos  " & _
'                      "WHERE  (Detalle_Facturas.Numero_Factura = '" & FacturaNumero & "') AND (Detalle_Facturas.Tipo_Factura = 'Factura') AND (Detalle_Facturas.Fecha_Factura = CONVERT(DATETIME, '" & Format(FechaFactura, "yyyy-mm-dd") & "', 102))"
                SQL = "SELECT  Productos.*, Detalle_Facturas.* FROM Detalle_Facturas INNER JOIN Productos ON Detalle_Facturas.Cod_Producto = Productos.Cod_Productos  " & _
                      "WHERE  (Detalle_Facturas.Numero_Factura = '" & FacturaNumero & "') AND (Detalle_Facturas.Tipo_Factura = '" & Tipo & "') "
                
                 
                  ArepFacturas.DataControl1.ConnectionString = ConexionFacturacion
                  ArepFacturas.DataControl1.Source = SQL
                
                   ArepFacturas.Logo.Picture = LoadPicture(RutaLogo)
                
                  ArepFacturas.LblEmpresa.Caption = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
                  ArepFacturas.LblEmpresa1.Caption = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
                  ArepFacturas.LblEmpresa2.Caption = "RUC " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
'                  ArepFacturas.LblCodigoCliente.Caption = CodigoCliente 'ArepEstadoCuenta.Field1.Text
'                  ArepFacturas.LblNombreCliente.Caption = NombreCliente 'ArepEstadoCuenta.Field2.Text
                  
                  ArepFacturas.Show 1
                  
               
'                     Set rpt = New ArepFacturas
'                     rpt.DataControl1.ConnectionString = ConexionFacturacion
'                     rpt.DataControl1.Source = SQL
'                     fPreview.RunReport rpt
'                     fPreview.Show 1
            
            
            End If
      
      Case "ArepAuxiliar"
      
       FechaTransaccion = LeeCadena(Link, 1)
       NumeroTransaccion = LeeCadena(Link, 2)
       
       
           
            ArepTransacciones.DataControl1.ConnectionString = ConexionReporte
                ArepTransacciones.LblNombre.Caption = "Comprobantes de Diario"
                SQL = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta AS DescripcionCuentas, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, " & _
                    "Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, " & _
                    "Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas, " & _
                    "Transacciones.NumeroMovimiento , Periodos.Periodo " & _
                    "FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo " & _
                    "WHERE  (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaTransaccion, "yyyymmdd") & "' And '" & Format(FechaTransaccion, "yyyymmdd") & "') AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ") " & _
                    "ORDER BY Transacciones.NTransaccion"
                    
                ArepTransacciones.DataControl1.Source = SQL
                
        
                
'                If UCase(Me.TxtFuente.Text) = UCase("Cheque") Then
'                    ArepTransacciones.LblNombre.Caption = "Comprobantes de Pago"
'                Else
'                    ArepTransacciones.LblNombre.Caption = "Comprobante de Diario"
'                End If
        
        
'                Dim rpt As Object
'                Dim fPreview As New FrmPreview
            
                 Set rpt = New ArepTransacciones
                 rpt.DataControl1.ConnectionString = ConexionReporte
                 rpt.DataControl1.Source = SQL
                 fPreview.RunReport rpt
                 fPreview.Show 1
                 QuienReporte = "ArepAuxiliar"

      Case "ArepEstadoCuenta"
      
            
            
            FacturaNumero = Link
            NumeroFact = Link
            If InStr(1, Link, "htm", vbTextCompare) = 0 And InStr(1, Link, "mailto", vbTextCompare) = 0 Then
            
                     If MDIPrimero.AdoConfiguracion.Recordset.RecordCount > 0 Then
                         If Not IsNull(MDIPrimero.AdoConfiguracion.Recordset!ConexionFacturacion) Then
                            ConexionFacturacion = MDIPrimero.AdoConfiguracion.Recordset!ConexionFacturacion
                         Else
                            ConexionFacturacion = ""
                         End If
                     End If
            
            
                FacturaNo = Link
                
                
'                FechaFactura = ArepEstadoCuenta.Field6.Text
                
                '---------------------------------------------------------------------------------------------------
                '-----------------------------------------CONSULTO LOS DATOS DE LA FACTURA ------------------------
                '---------------------------------------------------------------------------------------------------
'                SQL = "SELECT  * FROM  Facturas INNER JOIN Bodegas ON Facturas.Cod_Bodega = Bodegas.Cod_Bodega " & _
'                      "WHERE (Facturas.Numero_Factura = '" & FacturaNumero & "') AND (Facturas.Fecha_Factura = CONVERT(DATETIME, '" & Format(FechaFactura, "yyyy-mm-dd") & "', 102)) AND (Facturas.Tipo_Factura = N'Factura')"
                SQL = "SELECT  * FROM  Facturas INNER JOIN Bodegas ON Facturas.Cod_Bodega = Bodegas.Cod_Bodega " & _
                      "WHERE (Facturas.Numero_Factura = '" & FacturaNumero & "') AND (Facturas.Tipo_Factura = 'Factura')"
              
                MDIPrimero.AdoConsultaFacturacion.ConnectionString = ConexionFacturacion
                MDIPrimero.AdoConsultaFacturacion.RecordSource = SQL
                MDIPrimero.AdoConsultaFacturacion.Refresh
                If Not MDIPrimero.AdoConsultaFacturacion.Recordset.EOF Then
                  ArepFacturas.LblSubTotal.Caption = Format(MDIPrimero.AdoConsultaFacturacion.Recordset("SubTotal"), "##,##0.00")
                  ArepFacturas.LblIVA.Caption = Format(MDIPrimero.AdoConsultaFacturacion.Recordset("IVA"), "##,##0.00")
                  ArepFacturas.LblTotal.Caption = Format(MDIPrimero.AdoConsultaFacturacion.Recordset("SubTotal") + MDIPrimero.AdoConsultaFacturacion.Recordset("IVA"), "##,##0.00")
                  ArepFacturas.LblBodega.Caption = MDIPrimero.AdoConsultaFacturacion.Recordset("Cod_Bodega") & " " & MDIPrimero.AdoConsultaFacturacion.Recordset("Nombre_Bodega")
                  ArepFacturas.LblObservaciones.Caption = MDIPrimero.AdoConsultaFacturacion.Recordset("Observaciones")
                  ArepFacturas.LblCodigoCliente.Caption = MDIPrimero.AdoConsultaFacturacion.Recordset("Cod_Cliente")
                  ArepFacturas.LblNombreCliente.Caption = MDIPrimero.AdoConsultaFacturacion.Recordset("Nombre_Cliente")
                  ArepFacturas.LblNuestraRef.Caption = MDIPrimero.AdoConsultaFacturacion.Recordset("MonedaFactura")
                End If
                
            
                
                
'                SQL = "SELECT  Productos.*, Detalle_Facturas.* FROM Detalle_Facturas INNER JOIN Productos ON Detalle_Facturas.Cod_Producto = Productos.Cod_Productos  " & _
'                      "WHERE  (Detalle_Facturas.Numero_Factura = '" & FacturaNumero & "') AND (Detalle_Facturas.Tipo_Factura = 'Factura') AND (Detalle_Facturas.Fecha_Factura = CONVERT(DATETIME, '" & Format(FechaFactura, "yyyy-mm-dd") & "', 102))"
                SQL = "SELECT  Productos.*, Detalle_Facturas.* FROM Detalle_Facturas INNER JOIN Productos ON Detalle_Facturas.Cod_Producto = Productos.Cod_Productos  " & _
                      "WHERE  (Detalle_Facturas.Numero_Factura = '" & FacturaNumero & "') AND (Detalle_Facturas.Tipo_Factura = 'Factura') "
                
                 
                  ArepFacturas.DataControl1.ConnectionString = ConexionFacturacion
                  ArepFacturas.DataControl1.Source = SQL
                
                   ArepFacturas.Logo.Picture = LoadPicture(RutaLogo)
                
                  ArepFacturas.LblEmpresa.Caption = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
                  ArepFacturas.LblEmpresa1.Caption = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
                  ArepFacturas.LblEmpresa2.Caption = "RUC " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
'                  ArepFacturas.LblCodigoCliente.Caption = CodigoCliente 'ArepEstadoCuenta.Field1.Text
'                  ArepFacturas.LblNombreCliente.Caption = NombreCliente 'ArepEstadoCuenta.Field2.Text
                  
'                  ArepFacturas.Show 1
                  
               
                     Set rpt = New ArepFacturas
                     rpt.DataControl1.ConnectionString = ConexionFacturacion
                     rpt.DataControl1.Source = SQL
                     fPreview.RunReport rpt
                     fPreview.Show 1
                     QuienReporte = "ArepEstadoCuenta"
            
            End If
   End Select
   
   
   Exit Sub
TipoErrs:
   If err.Number = -2147467259 Then
     MsgBox "No Existe Conexion con el Modulo de Facturacion", vbCritical, "Zeus Contable"
   End If

   Exit Sub

End Sub

Private Sub ExportaExcel_Click()
Dim xls As New ActiveReportsExcelExport.ARExportExcel
Dim sFile As String
Dim bSave As Boolean

On Error GoTo TipoErrs
   
    Me.CommonDialog.Filter = "Formato Excel (*.xls)| *.xls"
    Me.CommonDialog.ShowSave
'    bSave = Dir(Me.CommonDialog.FileName + ".xls")
    
'    If bSave Then xls.FileName = sFile Else Exit Sub

    sFile = Me.CommonDialog.FileName
    xls.FileName = sFile
    
    If arv.Pages.Count > 0 Then
        xls.Export arv.Pages
    ElseIf Not arv.ReportSource Is Nothing Then
        If arv.ReportSource.Pages.Count > 0 Then
            xls.Export arv.ReportSource.Pages
        End If
    End If
    Set xls = Nothing
    MsgBox "Se ha Exportado el Archivo", vbExclamation, "Zeus Contabilidad"
    
    Exit Sub
TipoErrs:
 MsgBox err.Description
    
End Sub

Private Sub ExportaPDF_Click()
Dim pdf As New ActiveReportsPDFExport.ARExportPDF
Dim sFile As String
Dim bSave As Boolean

On Error GoTo TipoErrs

    Me.CommonDialog.Filter = "Portable Document Format (*.PDF)| *.PDF"
    Me.CommonDialog.ShowSave
    sFile = Me.CommonDialog.FileName
    
    
    pdf.FileName = sFile
    
    If arv.Pages.Count > 0 Then
        pdf.Export arv.Pages
    ElseIf Not arv.ReportSource Is Nothing Then
        If arv.ReportSource.Pages.Count > 0 Then
            pdf.Export arv.ReportSource.Pages
        End If
    End If
    
    Set pdf = Nothing
    MsgBox "Se ha Exportado el Archivo", vbExclamation, "Zeus Contabilidad"
    
Exit Sub
TipoErrs:
 MsgBox err.Description
End Sub

Public Sub RunReport(rpt As Object)
    Set arv.ReportSource = rpt
    
    arv.Zoom = 100
    Caption = rpt.Caption
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    arv.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
