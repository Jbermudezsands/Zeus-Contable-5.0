Attribute VB_Name = "Funciones"

Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst _
    As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 _
    As Long, ByVal un2 As Long) As Long
    Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Function Encrypt(Frase As String) As String
Dim Ilen As Integer, x As Integer
Dim sFrase As String, sCurrent As String, sNew As String
Ilen = Len(Frase)
For x = 1 To Ilen
    sCurrent = Mid$(Frase, x, 1)
    sNew = Chr$(Asc(sCurrent) + 110)
    sFrase = sFrase & sNew
Next
Encrypt = sFrase
End Function
'Public Function Inicio_Excel() As Boolean
'
'Dim i As Integer
'Dim J As Integer
'
'Set objExcel = New Excel.Application
'
'objExcel.Visible = True 'lo hacemos visible
'objExcel.SheetsInNewWorkbook = 1 'decimos cuantas hojas queremos en el nuevo documento
'objExcel.Workbooks.Add ' añadimos el objeto al workbook
'
'End Function


Public Function SaldoCuenta(Periodo As Double, Fecha As Date, Cuenta As String, KeyPresupuesto As Double) As Double
      Dim TipoMoneda As String, TipoCuenta As String, MontoTasa As Double
      
       TipoMoneda = "Córdobas"  'DtaCuentas.Recordset("TipoMoneda")
       TipoCuenta = "Gastos" 'DtaCuentas.Recordset("TipoCuenta")
       MontoTasa = BuscaTasaCambio(Fecha)
      
      
      '////////////////////////////////////////////////////////////////////////////////////////////////////////////
      '//////////////////////////CONSULTO LOS SALDOS ACUMULADO REAL DE LAS CUENTAS PARA PRESUPUESTO //////////////////
      '/////////////////////////////////////////////////////////////////////////////////////////////////
      MDIPrimero.AdoConsulta.RecordSource = "SELECT SUM(Debito * TCambio) AS MDebito, SUM(TCambio * Credito) AS MCredito From Transacciones WHERE   (FechaTransaccion  < CONVERT(DATETIME, '" & Format(Fecha, "YYYY-MM-DD") & "', 102)) GROUP BY FacturaNo HAVING (FacturaNo = '" & KeyPresupuesto & "')"
      MDIPrimero.AdoConsulta.Refresh

      If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
        If Not IsNull(MDIPrimero.AdoConsulta.Recordset("MDebito")) Then
          Debito = MDIPrimero.AdoConsulta.Recordset("MDebito")
        End If
      Else
         Debito = 0
      End If
      
      If Not IsNull(MDIPrimero.AdoConsulta.Recordset("MCredito")) Then
       Credito = MDIPrimero.AdoConsulta.Recordset("MCredito")
      Else
       Credito = 0
      End If
      
       Select Case TipoMoneda
         Case "Dólares"
            If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
             Saldo = (Debito - Credito)
            Else
             Saldo = (Credito - Debito)
            End If
         Case "Libras"
             If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                Saldo = (Debito - Credito)
             Else
                Saldo = (Credito - Debito)
             End If
         Case "Córdobas"
             If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                Saldo = (Debito - Credito)
             Else
                Saldo = (Credito - Debito)
             End If
       End Select


    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////BUSCO EL MONTO TOTAL PRESUPUESTADO ///////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////////////
    
    SaldoCuenta = Saldo

End Function



Public Sub ConvertirReporte(Fecha As Date)
  Dim TasaCambio As Double, Debe1 As Double, Debe2 As Double, Debe3 As Double, Haber1 As Double, Haber2 As Double, Haber3 As Double
  
  TasaCambio = BuscaTasaCambio(Fecha)
   '////////////////////////////////////////////BUSCO EL TOTAL DE ACTIVOS //////////////////////////////////////////////////
  MDIPrimero.AdoConsulta.RecordSource = "SELECT Reportes.* From Reportes ORDER BY Orden"
  MDIPrimero.AdoConsulta.Refresh
  Do While Not MDIPrimero.AdoConsulta.Recordset.EOF
    Debe1 = CDbl(MDIPrimero.AdoConsulta.Recordset("Debe1")) / TasaCambio
    Haber1 = CDbl(MDIPrimero.AdoConsulta.Recordset("Haber1")) / TasaCambio
    Debe2 = CDbl(MDIPrimero.AdoConsulta.Recordset("Debe2")) / TasaCambio
    Haber2 = CDbl(MDIPrimero.AdoConsulta.Recordset("Haber2")) / TasaCambio
    Debe3 = CDbl(MDIPrimero.AdoConsulta.Recordset("Debe3")) / TasaCambio
    Haber3 = CDbl(MDIPrimero.AdoConsulta.Recordset("Haber3")) / TasaCambio
    
    MDIPrimero.AdoConsulta.Recordset("Debe1") = Format(Debe1, "##0.00")
    MDIPrimero.AdoConsulta.Recordset("Haber1") = Format(Haber1, "##0.00")
    MDIPrimero.AdoConsulta.Recordset("Debe2") = Format(Debe2, "##0.00")
    MDIPrimero.AdoConsulta.Recordset("Haber2") = Format(Haber2, "##0.00")
    MDIPrimero.AdoConsulta.Recordset("Debe3") = Format(Debe3, "##0.00")
    MDIPrimero.AdoConsulta.Recordset("Haber3") = Format(Haber3, "##0.00")
    MDIPrimero.AdoConsulta.Recordset.Update
  
  
   MDIPrimero.AdoConsulta.Recordset.MoveNext
  Loop
End Sub




Public Sub AjusteDiferencial()
  Dim TotalActivo1 As Double, TotalPC1 As Double, Diferencial As Double, TotalActivo2 As Double, TotalActivo3 As Double, TotalPC2 As Double, TotalPC3 As Double
  Dim Orden As Double, OrdenPC As Double

  '////////////////////////////////////////////BUSCO EL TOTAL DE ACTIVOS //////////////////////////////////////////////////
  MDIPrimero.AdoConsulta.RecordSource = "SELECT Reportes.* From Reportes WHERE (KeyGrupo = 'A') AND (Descripcion LIKE N'%Total%')"
  MDIPrimero.AdoConsulta.Refresh
  If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
    TotalActivo1 = Format(MDIPrimero.AdoConsulta.Recordset("Haber1"), "##0.00")
    TotalActivo2 = Format(MDIPrimero.AdoConsulta.Recordset("Haber2"), "##0.00")
    TotalActivo3 = Format(MDIPrimero.AdoConsulta.Recordset("Haber3"), "##0.00")
  End If
  
  
  MDIPrimero.AdoConsulta.RecordSource = "SELECT Reportes.* From Reportes WHERE (KeyGrupo = 'PC') "
  MDIPrimero.AdoConsulta.Refresh
  If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
    TotalPC1 = Format(MDIPrimero.AdoConsulta.Recordset("Haber1"), "##0.00")
    TotalPC2 = Format(MDIPrimero.AdoConsulta.Recordset("Haber2"), "##0.00")
    TotalPC3 = Format(MDIPrimero.AdoConsulta.Recordset("Haber3"), "##0.00")
    OrdenPC = MDIPrimero.AdoConsulta.Recordset("Orden")
  End If
  
  
  Diferencial = TotalActivo1 - TotalPC1
  
  If Abs(Diferencial) < 0.02 And Abs(Diferencial) <> 0 Then
  
        
        Diferencial = Format(Diferencial, "##0.00")
  
        MDIPrimero.AdoConsulta.RecordSource = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden From Reportes ORDER BY Orden"
        MDIPrimero.AdoConsulta.Refresh
        If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
          MDIPrimero.AdoConsulta.Recordset.MoveLast
          Orden = MDIPrimero.AdoConsulta.Recordset("Orden") + 1
        End If
        
        
        MDIPrimero.AdoConsulta.Recordset.AddNew
           MDIPrimero.AdoConsulta.Recordset("Descripcion") = "Diferencial Cambiario"
           MDIPrimero.AdoConsulta.Recordset("KeyGrupo") = "DF"
           MDIPrimero.AdoConsulta.Recordset("Haber1") = Diferencial
           MDIPrimero.AdoConsulta.Recordset("Nivel") = 1
           MDIPrimero.AdoConsulta.Recordset("Orden") = OrdenPC
        MDIPrimero.AdoConsulta.Recordset.Update
        
        MDIPrimero.AdoConsulta.RecordSource = "SELECT Reportes.* From Reportes WHERE (KeyGrupo = 'PC') "
        MDIPrimero.AdoConsulta.Refresh
        If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
          MDIPrimero.AdoConsulta.Recordset("Haber1") = MDIPrimero.AdoConsulta.Recordset("Haber1") + Diferencial
          MDIPrimero.AdoConsulta.Recordset("Haber2") = MDIPrimero.AdoConsulta.Recordset("Haber2") + Diferencial
          MDIPrimero.AdoConsulta.Recordset("Haber3") = MDIPrimero.AdoConsulta.Recordset("Haber3") + Diferencial
          MDIPrimero.AdoConsulta.Recordset("Orden") = Orden
          MDIPrimero.AdoConsulta.Recordset.Update
        End If
    
    
  End If

End Sub



Function Decrypt(Frase As String) As String
Dim Ilen As Integer, x As Integer
Dim sFrase As String, sCurrent As String, sNew As String
Ilen = Len(Frase)
For x = 1 To Ilen
    sCurrent = Mid$(Frase, x, 1)
    sNew = Chr$(Asc(sCurrent) - 110)
    sFrase = sFrase & sNew
Next
Decrypt = sFrase
End Function
 Public Function MostrarBotonDpto(CodCuenta As String) As Boolean
  '///////////////////////////////////////////////////////////////////////////////////////////////////////
  '/////////////////////////////////BUSCO EL INDICE TRANSACCIONES ////////////////////////////////////////
  '///////////////////////////////////////////////////////////////////////////////////////////////////////
  MDIPrimero.AdoConsulta.RecordSource = "SELECT  * From Cuentas WHERE (CodCuentas = '" & CodCuenta & "') "
  MDIPrimero.AdoConsulta.Refresh
  If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
    Select Case MDIPrimero.AdoConsulta.Recordset("TipoCuenta")
       Case "Ingresos - Ventas"
          MostrarBotonDpto = True
       Case "Costos"
          MostrarBotonDpto = True
       Case "Gastos"
          MostrarBotonDpto = True
    End Select
  Else
      MostrarBotonDpto = False
      Exit Function
  End If


 End Function
Public Function ExisteDpto(CodDepartamento As String) As Boolean
  '///////////////////////////////////////////////////////////////////////////////////////////////////////
  '/////////////////////////////////BUSCO EL INDICE TRANSACCIONES ////////////////////////////////////////
  '///////////////////////////////////////////////////////////////////////////////////////////////////////
  MDIPrimero.AdoConsulta.RecordSource = "SELECT  * From GrupoCuentas WHERE  (CodGrupo = '" & CodDepartamento & "')"
  MDIPrimero.AdoConsulta.Refresh
  If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
   ExisteDpto = True
  Else
    ExisteDpto = False
  End If

End Function

Public Function BuscaDpto(CodDepartamento As String) As String
  '///////////////////////////////////////////////////////////////////////////////////////////////////////
  '/////////////////////////////////BUSCO EL INDICE TRANSACCIONES ////////////////////////////////////////
  '///////////////////////////////////////////////////////////////////////////////////////////////////////
  MDIPrimero.AdoConsulta.RecordSource = "SELECT  * From GrupoCuentas WHERE  (CodGrupo = '" & CodDepartamento & "')"
  MDIPrimero.AdoConsulta.Refresh
  If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
   BuscaDpto = MDIPrimero.AdoConsulta.Recordset("DescripcionGrupo")
  Else
    BuscaDpto = "Sin Departamento"
  End If

End Function
    
Public Function Inicio_Excel() As Boolean
Dim i As Integer
Dim J As Integer

Set objExcel = New Excel.Application

objExcel.Visible = True 'lo hacemos visible
objExcel.SheetsInNewWorkbook = 1 'decimos cuantas hojas queremos en el nuevo documento
objExcel.Workbooks.Add ' añadimos el objeto al workbook

End Function

Public Function ConvertirMovimiento(NumeroMovimiento As Double, Fecha As Date, MonedaConvertir As String)
  Dim TipoMoneda As String, TasaCambio As Double
  '///////////////////////////////////////////////////////////////////////////////////////////////////////
  '/////////////////////////////////BUSCO EL INDICE TRANSACCIONES ////////////////////////////////////////
  '///////////////////////////////////////////////////////////////////////////////////////////////////////
  MDIPrimero.AdoConsulta.RecordSource = "SELECT * From IndiceTransaccion WHERE (FechaTransaccion = CONVERT(DATETIME, '" & Format(Fecha, "YYYY-MM-DD") & " ', 102)) AND (NumeroMovimiento = " & NumeroMovimiento & ")"
  MDIPrimero.AdoConsulta.Refresh
  If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
    If MDIPrimero.AdoConsulta.Recordset("TipoMoneda") = MonedaConvertir Then
      Exit Function
    End If
  Else
      Exit Function
  End If
  
  '////////////////////////////////////////EDITO EL INDICE DE TRANSACCION/////////////////////////////////////////
  MDIPrimero.AdoConsulta.Refresh
  If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
    TipoMoneda = MDIPrimero.AdoConsulta.Recordset("TipoMoneda")
    MDIPrimero.AdoConsulta.Recordset("TipoMoneda") = MonedaConvertir
    MDIPrimero.AdoConsulta.Recordset.Update
  End If
  
  TasaCambio = BuscaTasaCambio(Fecha)
  
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '//////////////////////////////CAMBIO LOS MONTOS DE LA TRANSACCION///////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    MDIPrimero.AdoConsulta.RecordSource = "SELECT  * From Transacciones WHERE (FechaTransaccion = CONVERT(DATETIME, '" & Format(Fecha, "YYYY-MM-DD") & " ', 102)) AND (NumeroMovimiento = " & NumeroMovimiento & ")"
    MDIPrimero.AdoConsulta.Refresh
    Do While Not MDIPrimero.AdoConsulta.Recordset.EOF
      
     If TipoMoneda = "Córdobas" Then
       MDIPrimero.AdoConsulta.Recordset("TCambio") = TasaCambio
       MDIPrimero.AdoConsulta.Recordset("Debito") = MDIPrimero.AdoConsulta.Recordset("Debito") / TasaCambio
       MDIPrimero.AdoConsulta.Recordset("Credito") = MDIPrimero.AdoConsulta.Recordset("Credito") / TasaCambio
       MDIPrimero.AdoConsulta.Recordset.Update
     Else
       MDIPrimero.AdoConsulta.Recordset("TCambio") = 1
       MDIPrimero.AdoConsulta.Recordset("Debito") = MDIPrimero.AdoConsulta.Recordset("Debito") * TasaCambio
       MDIPrimero.AdoConsulta.Recordset("Credito") = MDIPrimero.AdoConsulta.Recordset("Credito") * TasaCambio
       MDIPrimero.AdoConsulta.Recordset.Update
     End If
     
    
     MDIPrimero.AdoConsulta.Recordset.MoveNext
    Loop
  
  




End Function


Public Function LeeCadena(cadena As String, numero As Double) As String
  Dim i As Double
  Dim A() As String, Can As String, J As Double
  
'  ReDim A(Len(Cadena))
   ReDim A(5)

    J = 1
    For i = 1 To Len(cadena)
     If Mid(cadena, i, 1) = ";" Then
      J = J + 1
      i = i + 1
      Can = ""
     End If
     
        Can = Can & Mid(cadena, i, 1)
        A(J) = Can
         

    Next
    
    LeeCadena = A(numero)
    
End Function


Public Function TRUNC(Value As Double, escala As Double)
 Dim posicion As Double, Valor As Control, Dec As Double, i As Double
 Dec = 1
 For i = 1 To escala
    Dec = Dec * 10
 Next
 
 Control = Value
'  TRUNC = Int(Value * 1000) / 1000
   TRUNC = Int(Control * Dec) / Dec
   
End Function
Public Function BuscaTasaCambio(FechaTasa As Date) As Double
      MDIPrimero.AdoConsulta.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas WHERE (FechaTasas = '" & Format(FechaTasa, "yyyymmdd") & "')"
      MDIPrimero.AdoConsulta.Refresh
      If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
         BuscaTasaCambio = MDIPrimero.AdoConsulta.Recordset("MontoCordobas")
      End If
      
      
      
End Function

Public Function BuscaTasaCambioFacturacion(FechaTasa As Date, ConexionFactura As String) As Double
      MDIPrimero.AdoConsultaFacturacion.ConnectionString = ConexionFactura
      MDIPrimero.AdoConsultaFacturacion.RecordSource = "SELECT   FechaTasa, MontoTasa From TasaCambio WHERE (FechaTasa = CONVERT(DATETIME, '" & Format(FechaTasa, "yyyymmdd") & "', 102))"
      MDIPrimero.AdoConsultaFacturacion.Refresh
      If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
         BuscaTasaCambioFacturacion = MDIPrimero.AdoConsultaFacturacion.Recordset("MontoTasa")
      Else
         BuscaTasaCambioFacturacion = 0
      End If
End Function


Public Function CalcularCostoPromedio(CodigoProducto As String, Conexion As String) As Double
 Dim MonedaCompra As String, PrecioUnitario As Double, Cantidad As Double, TotalImporte As Double, TasaCambio As Double, FechaCompra As Date
 Dim PrecioCostoDolar As Double, PrecioCosto As Double, Importe As Double

          MDIPrimero.AdoConsultaFacturacion.ConnectionString = Conexion
          MDIPrimero.AdoConsultaFacturacion.RecordSource = "SELECT MAX(Detalle_Compras.Numero_Compra) AS Numero_Compra, MAX(Detalle_Compras.Fecha_Compra) AS Fecha_Compra, Detalle_Compras.Tipo_Compra, Detalle_Compras.Cod_Producto, SUM(Detalle_Compras.Cantidad) AS Cantidad, SUM(Detalle_Compras.Precio_Unitario) AS Precio_Unitario, SUM(Detalle_Compras.Descuento) AS Descuento, SUM(Detalle_Compras.Precio_Neto) AS Precio_Neto, Compras.MonedaCompra, SUM(Detalle_Compras.Precio_Unitario * Detalle_Compras.Cantidad) AS Importe FROM Detalle_Compras INNER JOIN Compras ON Detalle_Compras.Numero_Compra = Compras.Numero_Compra AND Detalle_Compras.Fecha_Compra = Compras.Fecha_Compra AND Detalle_Compras.Tipo_Compra = Compras.Tipo_Compra GROUP BY Detalle_Compras.Tipo_Compra, Detalle_Compras.Cod_Producto, Compras.MonedaCompra HAVING (Detalle_Compras.Cod_Producto = '" & CodigoProducto & "') AND (Detalle_Compras.Tipo_Compra = 'Mercancia Recibida')"
          MDIPrimero.AdoConsultaFacturacion.Refresh
          Do While Not MDIPrimero.AdoConsultaFacturacion.Recordset.EOF
                  MonedaCompra = MDIPrimero.AdoConsultaFacturacion.Recordset("MonedaCompra")
                  If MonedaCompra = "Cordobas" Then
                     PrecioUnitario = Trim(MDIPrimero.AdoConsultaFacturacion.Recordset("Precio_Unitario"))
                     Cantidad = Trim(MDIPrimero.AdoConsultaFacturacion.Recordset("Cantidad") + Cantidad)
                     Importe = Trim(MDIPrimero.AdoConsultaFacturacion.Recordset("Importe"))
                     TotalImporte = TotalImporte + Importe
                     FechaCompra = Trim(MDIPrimero.AdoConsultaFacturacion.Recordset("Fecha_Compra"))
                     TasaCambio = BuscaTasaCambio(FechaCompra)
                     If TasaCambio = 0 Then
                        MsgBox "TASA DE CAMBIO CERO", vbApplicationModal, "Zeus Contabilidad "
                        TasaCambio = 1
                     Else
                        PrecioCostoDolar = (TotalImporte / TasaCambio)
                     End If
                  Else
                     PrecioUnitario = Trim(MDIPrimero.AdoConsultaFacturacion.Recordset("Precio_Unitario"))
                     Cantidad = Trim(MDIPrimero.AdoConsultaFacturacion.Recordset("Cantidad") + Cantidad)
                     Importe = Trim(MDIPrimero.AdoConsultaFacturacion.Recordset("Importe"))
                     FechaCompra = Trim(MDIPrimero.AdoConsultaFacturacion.Recordset("Fecha_Compra"))
                     TasaCambio = BuscaTasaCambio(FechaCompra)
                     If TasaCambio = 0 Then
                        MsgBox "TASA DE CAMBIO CERO " & FechaCompra, vbApplicationModal, "Zeus Contabilidad "
                     Else
                        TotalImporte = (Importe * TasaCambio) + TotalImporte
                     End If
                  End If

               MDIPrimero.AdoConsultaFacturacion.Recordset.MoveNext
           Loop
           
                    If Cantidad <> 0 Then
                        PrecioCosto = Format(TotalImporte / Cantidad, "##,##0.00")
                        PrecioCostoDolar = Format((PrecioCosto / TasaCambio), "##,##0.00")
                    Else
                        PrecioCosto = PrecioUnitario
                        If PrecioCosto <> 0 Then
                            PrecioCostoDolar = Format((PrecioCosto / TasaCambio), "##,##0.00")
                        Else
                            PrecioCostoDolar = 0
                        End If
                    End If

    CalcularCostoPromedio = PrecioCosto

End Function


Public Function ValidarCuentas(CodigoCuentas As String) As Boolean
           
           Sql = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo " & _
                 "From Cuentas WHERE (CodCuentas = '" & CodigoCuentas & "')"
           MDIPrimero.AdoConsulta.RecordSource = Sql
           MDIPrimero.AdoConsulta.Refresh
           If MDIPrimero.AdoConsulta.Recordset.EOF Then
              ValidarCuentas = False
           Else
              ValidarCuentas = True
           End If

End Function


Public Function MesLetras(Meses As Double) As String
  Select Case Meses
      Case 1: MesLetras = "Enero"
      Case 2: MesLetras = "Febrero"
      Case 3: MesLetras = "Marzo"
      Case 4: MesLetras = "Abril"
      Case 5: MesLetras = "Mayo"
      Case 6: MesLetras = "Junio"
      Case 7: MesLetras = "Julio"
      Case 8: MesLetras = "Julio"
      Case 9: MesLetras = "Agosto"
      Case 10: MesLetras = "Septiembre"
      Case 11: MesLetras = "Octubre"
      Case 12: MesLetras = "Noviembre"
      
     
  End Select

End Function


Sub HayUtilidadBruta()
 HayUtiBruta = True
        With FrmReportes.DtaReportes
            .RecordSource = "select * from reportes where descripcion like '%" & MDIPrimero.AdoConfiguracion.Recordset!OtrosINgresos & "%' order by orden"
            .Refresh
            If .Recordset.RecordCount = 0 Then HayUtiBruta = False
            .RecordSource = "select * from reportes where descripcion like '%" & MDIPrimero.AdoConfiguracion.Recordset!Costos & "%' order by orden"
            .Refresh
            If .Recordset.RecordCount = 0 Then HayUtiBruta = False
            .RecordSource = "select * from reportes where descripcion like '%" & MDIPrimero.AdoConfiguracion.Recordset!Ingresos & "%' order by orden"
            .Refresh
            If .Recordset.RecordCount = 0 Then HayUtiBruta = False
            .RecordSource = "select * from reportes where descripcion like '%" & MDIPrimero.AdoConfiguracion.Recordset!CostosOperativos & "%' order by orden"
            .Refresh
            If .Recordset.RecordCount = 0 Then HayUtiBruta = False
        End With
        
End Sub
Public Sub Demo(FechaSalida As Date)
 Dim Horas As Double, HorasEncriptadas As String
 Dim cadena As String, Cadena2 As String, Cadena3 As String  ' Cadena3 its an decrypted value from cadena2
 Dim result As Double
 
' Cadena = Encrypt("Siempre")
 cadena = "ABC"
 
    '---------------------------LEER DATOS EMPRESA ------------------------
    MDIPrimero.AdoConsulta.ConnectionString = Conexion
    MDIPrimero.AdoConsulta.RecordSource = "DatosEmpresa"
    MDIPrimero.AdoConsulta.Refresh
    
    If Not IsNull(MDIPrimero.AdoConsulta.Recordset("Valor")) Then
     If Not MDIPrimero.AdoConsulta.Recordset("Valor") = "" Then
      Cadena2 = MDIPrimero.AdoConsulta.Recordset("Valor")
     End If
    End If
    
    Cadena2 = Decrypt(Cadena2)
    Cadena3 = Mid(Cadena2, 1, 3)
    If Cadena3 = "ABC" Then
   
      Horas = DateDiff("n", FechaIngreso, FechaSalida)
      result = CDbl(Mid(Cadena2, 4, 9)) + Horas
     
      cadena = cadena & result
      HorasEncriptadas = Encrypt(cadena)
'      Cadena2 = Decrypt(HorasEncriptadas)
      
      MDIPrimero.AdoConsulta.Recordset("Valor") = HorasEncriptadas
      MDIPrimero.AdoConsulta.Recordset.Update
    
    ElseIf Not Cadena2 = "Siempre" Then
    
      cadena = "ABC3600"
      HorasEncriptadas = Encrypt(cadena)
      MDIPrimero.AdoConsulta.Recordset("Valor") = HorasEncriptadas
      MDIPrimero.AdoConsulta.Recordset.Update
    
    End If
    
    
    




End Sub

Public Sub KillProcess(ByVal processName As String)
On Error GoTo errHandler
Dim oWMI
Dim ret
Dim sService
Dim oWMIServices
Dim oWMIService
Dim oServices
Dim oService
Dim servicename
Set oWMI = GetObject("winmgmts:")
Set oServices = oWMI.InstancesOf("win32_process")
For Each oService In oServices

servicename = LCase(Trim(CStr(oService.Name) & ""))

If InStr(1, servicename, LCase(processName), vbTextCompare) > 0 Then
ret = oService.Terminate
End If

Next

Set oServices = Nothing
Set oWMI = Nothing

errHandler:
err.Clear
End Sub


Public Sub SaldosPersonalizados(TipoCuenta As String)
    
    ResultadoPersonalizado = 0
    ResultadoPersonalizadoPeriodo = 0
    FrmReportes.DtaConsulta.RecordSource = "SELECT MAX(Descripcion) AS Descripcion, SUM(Debe1) AS Debe1, SUM(Haber1) AS Haber1, SUM(Debe2) AS Debe2, SUM(Haber2) AS Haber2, SUM(Debe3) " & _
                      "AS Debe3, SUM(Haber3) AS Haber3, SUM(Debe3) - SUM(Haber3) AS Saldo, SUM(Debe2) - SUM(Haber2) AS SaldoPeriodo, MAX(KeyGrupo) AS KeyGrupo, MAX(KeyGrupoSuperior) " & _
                      "AS KeyGrupoSuperior, MAX(KeyGrupoCuenta) AS KeyGrupoCuenta, MAX(Nivel) AS Nivel, MAX(Orden) AS Orden, MAX(CodCuentas) AS CodCuentas, " & _
                      "Ubicacion As Ubicacion From Reportes GROUP BY Ubicacion HAVING (Ubicacion = '" & TipoCuenta & "') "
     FrmReportes.DtaConsulta.Refresh
     If Not FrmReportes.DtaConsulta.Recordset.EOF Then
        ResultadoPersonalizado = FrmReportes.DtaConsulta.Recordset("Saldo")
        ResultadoPersonalizadoPeriodo = FrmReportes.DtaConsulta.Recordset("SaldoPeriodo")
     End If


End Sub
Function FechaPeriodo(Periodo As Double) As String
            NumeroPeriodo1 = FrmReportes.CmbIni.Text
            NumeroPeriodo2 = FrmReportes.CmbFin.Text
            
            If FrmReportes.Option8 = True Then
             NumeroTabla = 1
            ElseIf FrmReportes.Option7 = True Then
              NumeroTabla = 2
            ElseIf FrmReportes.Option6 = True Then
              NumeroTabla = 3
            End If
            
            '///////////////////////////////////////BUSCO EL AÑO SELECCIONADO/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
              FrmReportes.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) Between " & Periodo & " And " & Periodo & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
              FrmReportes.DtaConsulta.Refresh
              If FrmReportes.DtaConsulta.Recordset.RecordCount = 0 Then
                MsgBox "Hay un problema con los Periodos, por Favor Revise los Periodos para buscar alguna anomalía", vbCritical
                Exit Function
              End If
               FrmReportes.DtaConsulta.Recordset.MoveLast
               i = FrmReportes.DtaConsulta.Recordset.RecordCount
               FrmReportes.DtaConsulta.Recordset.MoveFirst
              
              If Not FrmReportes.DtaConsulta.Recordset.EOF Then
        
        
                If i = 1 Then
                  FechaIni = "01/" & Month(FrmReportes.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
                  NumFecha1 = FechaIni
                  FechaFin = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
                  FechaPeriodo = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
                  NumFecha2 = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
                Else
        
                 If NumeroPeriodo1 = FrmReportes.DtaConsulta.Recordset("Periodo") Then
                  FechaIni = "01/" & Month(FrmReportes.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
                  NumFecha1 = FechaIni
                  FechaFin = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
                  FechaPeriodo = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
                 ElseIf NumeroPeriodo2 = FrmReportes.DtaConsulta.Recordset("Periodo") Then
                  FechaFin = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
                  NumFecha2 = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
                  FechaPeriodo = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
                 End If
                End If
                
               End If

End Function
Function FechaPeriodoIni(Periodo As Double) As String
            NumeroPeriodo1 = FrmReportes.CmbIni.Text
            NumeroPeriodo2 = FrmReportes.CmbFin.Text
            
            If FrmReportes.Option8 = True Then
             NumeroTabla = 1
            ElseIf FrmReportes.Option7 = True Then
              NumeroTabla = 2
            ElseIf FrmReportes.Option6 = True Then
              NumeroTabla = 3
            End If
            
            '///////////////////////////////////////BUSCO EL AÑO SELECCIONADO/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
              FrmReportes.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) Between " & Periodo & " And " & Periodo & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
              FrmReportes.DtaConsulta.Refresh
              If FrmReportes.DtaConsulta.Recordset.RecordCount = 0 Then
                MsgBox "Hay un problema con los Periodos, por Favor Revise los Periodos para buscar alguna anomalía", vbCritical
                Exit Function
              End If
               FrmReportes.DtaConsulta.Recordset.MoveLast
               i = FrmReportes.DtaConsulta.Recordset.RecordCount
               FrmReportes.DtaConsulta.Recordset.MoveFirst
              
              If Not FrmReportes.DtaConsulta.Recordset.EOF Then
        
        
                If i = 1 Then
                  FechaIni = "01/" & Month(FrmReportes.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
                  NumFecha1 = FechaIni
                  FechaFin = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
                  FechaPeriodoIni = FechaIni
                  NumFecha2 = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
                Else
        
                 If NumeroPeriodo1 = FrmReportes.DtaConsulta.Recordset("Periodo") Then
                  FechaIni = "01/" & Month(FrmReportes.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
                  NumFecha1 = FechaIni
                  FechaFin = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
                  FechaPeriodoIni = FechaIni
                 ElseIf NumeroPeriodo2 = FrmReportes.DtaConsulta.Recordset("Periodo") Then
                  FechaFin = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
                  NumFecha2 = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
                  FechaPeriodoIni = FechaIni
                 End If
                End If
                
               End If

End Function

Function SaldosRazonesDebitos(Fecha2 As String, TipoCuenta As Variant) As Double
 
             ResultadoPersonalizado = 0
             ResultadoPersonalizadoPeriodo = 0
             
             
             FrmReportes.DtaConsulta.RecordSource = "SELECT MAX(Cuentas.CodCuentas) AS Expr3, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito,SUM(Transacciones.Debito * Transacciones.TCambio - Transacciones.TCambio * Transacciones.Credito) AS Total, MAX(Cuentas.DescripcionCuentas) AS Descripcion, MAX(Cuentas.TipoMoneda) AS TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas " & _
                                                    "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) AND (Cuentas.TipoCuenta = '" & TipoCuenta & "') "
             
'              FrmReportes.DtaConsulta.RecordSource = "SELECT MAX(Cuentas.CodCuentas) AS Expr3, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio - Transacciones.TCambio * Transacciones.Credito) AS Total, MAX(Cuentas.DescripcionCuentas) AS Descripcion, MAX(Cuentas.TipoMoneda) AS TipoMoneda  " & _
'                                            "FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
'                                            "WHERE  (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Cuentas.TipoCuenta,Cuentas.SubDivicion " & _
'                                            "HAVING (Cuentas.TipoCuenta = '" & TipoCuenta & "') "
                FrmReportes.DtaConsulta.Refresh
                If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                  If Not IsNull(FrmReportes.DtaConsulta.Recordset("Total")) Then
                   ResultadoPersonalizado = FrmReportes.DtaConsulta.Recordset("Total")
                   SaldosRazonesDebitos = FrmReportes.DtaConsulta.Recordset("Total")
                  End If
                End If
End Function

Function SaldosRazonesDebitosFijo(Fecha2 As String, TipoCuenta As Variant, TipoGasto As String) As Double
 
             ResultadoPersonalizado = 0
             ResultadoPersonalizadoPeriodo = 0
             
             
             
               FrmReportes.DtaConsulta.RecordSource = "SELECT MAX(Cuentas.CodCuentas) AS Expr3, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.TCambio * Transacciones.Credito) AS Total, MAX(Cuentas.DescripcionCuentas) AS Descripcion, MAX(Cuentas.TipoMoneda) AS TipoMoneda  " & _
                                            "FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
                                            "WHERE  (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Cuentas.TipoCuenta, Cuentas.SubDivicion " & _
                                            "HAVING (Cuentas.TipoCuenta = '" & TipoCuenta & "') AND (Cuentas.SubDivicion = '" & TipoGasto & "')"
                FrmReportes.DtaConsulta.Refresh
                If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                   ResultadoPersonalizado = FrmReportes.DtaConsulta.Recordset("Total")
                   SaldosRazonesDebitosFijo = FrmReportes.DtaConsulta.Recordset("Total")
'                   ResultadoPersonalizadoPeriodo = FrmReportes.DtaConsulta.Recordset("SaldoPeriodo")
                End If
End Function
Function SaldoPeriodoDebitoFijo(Fecha1 As String, Fecha2 As String, TipoCuenta As Variant, TipoGrupo As String) As Double
 
             ResultadoPersonalizado = 0
             ResultadoPersonalizadoPeriodo = 0
              FrmReportes.DtaConsulta.RecordSource = "SELECT MAX(Cuentas.CodCuentas) AS Expr3, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.TCambio * Transacciones.Credito) AS Total, MAX(Cuentas.DescripcionCuentas) AS Descripcion, MAX(Cuentas.TipoMoneda) AS TipoMoneda  " & _
                                            "FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
                                            "WHERE   (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Cuentas.TipoCuenta, Cuentas.SubDivicion " & _
                                            "HAVING (Cuentas.TipoCuenta = '" & TipoCuenta & "') AND (Cuentas.SubDivicion = '" & TipoGrupo & "')"
                FrmReportes.DtaConsulta.Refresh
                If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                   ResultadoPersonalizado = FrmReportes.DtaConsulta.Recordset("Total")
                   SaldoPeriodoDebitoFijo = FrmReportes.DtaConsulta.Recordset("Total")
'                   ResultadoPersonalizadoPeriodo = FrmReportes.DtaConsulta.Recordset("SaldoPeriodo")
                End If
End Function

Function SaldoPeriodoDebito(Fecha1 As String, Fecha2 As String, TipoCuenta As Variant) As Double
 
             ResultadoPersonalizado = 0
             ResultadoPersonalizadoPeriodo = 0
              FrmReportes.DtaConsulta.RecordSource = "SELECT MAX(Cuentas.CodCuentas) AS Expr3, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.TCambio * Transacciones.Credito) AS Total, MAX(Cuentas.DescripcionCuentas) AS Descripcion, MAX(Cuentas.TipoMoneda) AS TipoMoneda  " & _
                                            "FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
                                            "WHERE   (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Cuentas.TipoCuenta " & _
                                            "HAVING (Cuentas.TipoCuenta = '" & TipoCuenta & "') "
                FrmReportes.DtaConsulta.Refresh
                If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                   ResultadoPersonalizado = FrmReportes.DtaConsulta.Recordset("Total")
                   SaldoPeriodoDebito = FrmReportes.DtaConsulta.Recordset("Total")
'                   ResultadoPersonalizadoPeriodo = FrmReportes.DtaConsulta.Recordset("SaldoPeriodo")
                End If
End Function
Function SaldoPeriodoCuentaDebito(Fecha1 As String, Fecha2 As String, Cuenta As Variant) As Double
 
             ResultadoPersonalizado = 0
             ResultadoPersonalizadoPeriodo = 0
              MDIPrimero.AdoConsulta.RecordSource = "SELECT  Cuentas.CodCuentas AS CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.TCambio * Transacciones.Credito) AS Total, MAX(Cuentas.DescripcionCuentas) AS Descripcion, MAX(Cuentas.TipoMoneda) AS TipoMoneda " & _
                                                    "FROM   Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas " & _
                                                    "WHERE  (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Cuentas.TipoCuenta, Cuentas.CodCuentas  HAVING (Cuentas.CodCuentas = '" & Cuenta & "')"
                MDIPrimero.AdoConsulta.Refresh
                If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
                   ResultadoPersonalizado = MDIPrimero.AdoConsulta.Recordset("Total")
                   SaldoPeriodoCuentaDebito = MDIPrimero.AdoConsulta.Recordset("Total")

                End If
End Function




Function SaldoPeriodoCredito(Fecha1 As String, Fecha2 As String, TipoCuenta As Variant) As Double
 
             ResultadoPersonalizado = 0
             ResultadoPersonalizadoPeriodo = 0
              FrmReportes.DtaConsulta.RecordSource = "SELECT MAX(Cuentas.CodCuentas) AS Expr3, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.TCambio * Transacciones.Credito)-SUM(Transacciones.Debito * Transacciones.TCambio) AS Total, MAX(Cuentas.DescripcionCuentas) AS Descripcion, MAX(Cuentas.TipoMoneda) AS TipoMoneda  " & _
                                            "FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
                                            "WHERE  (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Cuentas.TipoCuenta " & _
                                            "HAVING (Cuentas.TipoCuenta = '" & TipoCuenta & "') "
                FrmReportes.DtaConsulta.Refresh
                If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                   ResultadoPersonalizado = FrmReportes.DtaConsulta.Recordset("Total")
                   SaldoPeriodoCredito = FrmReportes.DtaConsulta.Recordset("Total")
'                   ResultadoPersonalizadoPeriodo = FrmReportes.DtaConsulta.Recordset("SaldoPeriodo")
                End If
End Function




Function SaldosRazonesCreditos(Fecha2 As String, TipoCuenta As Variant) As Double
 
             ResultadoPersonalizado = 0
             ResultadoPersonalizadoPeriodo = 0
              FrmReportes.DtaConsulta.RecordSource = "SELECT MAX(Cuentas.CodCuentas) AS Expr3, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.TCambio * Transacciones.Credito)-SUM(Transacciones.Debito * Transacciones.TCambio) AS Total, MAX(Cuentas.DescripcionCuentas) AS Descripcion, MAX(Cuentas.TipoMoneda) AS TipoMoneda  " & _
                                            "FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
                                            "WHERE  (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Cuentas.TipoCuenta " & _
                                            "HAVING (Cuentas.TipoCuenta = '" & TipoCuenta & "') "
                FrmReportes.DtaConsulta.Refresh
                If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                   ResultadoPersonalizado = FrmReportes.DtaConsulta.Recordset("Total")
                   SaldosRazonesCreditos = FrmReportes.DtaConsulta.Recordset("Total")
'                   ResultadoPersonalizadoPeriodo = FrmReportes.DtaConsulta.Recordset("SaldoPeriodo")
                End If
End Function

Function SaldosRazonesUbicacionCredito(Fecha2 As String, TipoCuenta As Variant) As Double
 
             ResultadoPersonalizado = 0
             ResultadoPersonalizadoPeriodo = 0
              FrmReportes.DtaConsulta.RecordSource = "SELECT MAX(Cuentas.CodCuentas) AS Expr3, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.TCambio * Transacciones.Credito)-SUM(Transacciones.Debito * Transacciones.TCambio) AS Total, MAX(Cuentas.DescripcionCuentas) AS Descripcion, MAX(Cuentas.TipoMoneda) AS TipoMoneda, Cuentas.UbicacionReporte  " & _
                                                     "FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas " & _
                                                     "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
                                                     "GROUP BY Cuentas.UbicacionReporte " & _
                                                     "HAVING      (Cuentas.UbicacionReporte = '" & TipoCuenta & "') "
                FrmReportes.DtaConsulta.Refresh
                If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                   ResultadoPersonalizado = FrmReportes.DtaConsulta.Recordset("Total")
                   SaldosRazonesUbicacionCredito = FrmReportes.DtaConsulta.Recordset("Total")
'                   ResultadoPersonalizadoPeriodo = FrmReportes.DtaConsulta.Recordset("SaldoPeriodo")
                End If
End Function


Public Sub Utilidadbruta()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'validar que se encuentrenn todas las cuentas para encontrar la utilidad bruta
        
       HayUtilidadBruta 've si se puede poner la utilidad bruta
    If HayUtiBruta Then
        With FrmReportes.DtaReportes
        'encuentra el último registro
        .RecordSource = "select * from reportes order by orden"
        .Refresh
        .Recordset.MoveLast
        UltimoOrden = .Recordset!Orden
        
        'encuentra el primer registro de otros ingresos
'        .RecordSource = "select * from reportes where descripcion like '%Otros Ingresos%' order by orden"
        .RecordSource = "select * from reportes where descripcion like '%" & MDIPrimero.AdoConfiguracion.Recordset!OtrosINgresos & "%' order by orden"
        .Refresh

        
         PrimReg = .Recordset!Orden


        'encuentra el último registro de otros ingresos
        '''''''''''''''''''''''''''''''''''''''
        'encuentra el primer registro de costos operativos y de este saldra el ultreg ya que movere todo lo que esta despues de total ingresoso operativos y antes de costos operativos 270406
        .RecordSource = "select * from reportes where descripcion like '%" & MDIPrimero.AdoConfiguracion.Recordset!Costos & "%' order by orden"
        .Refresh
'        .Recordset.MoveLast 'lo que quiero es el primer registro de costos operativos no el ultimo
        UltReg = .Recordset!Orden
        '''''''''''''''''''''''''''''''''''''''
        
        'encuentra la cantidad de registros de otros ingresos
        RegIngresos = UltReg - PrimReg
        
        'cambia el orden del último registro(utilidad o pérdida del ejercicio) sumándole la cantidad de registros de otros ingresos más uno
        .RecordSource = "select * from reportes order by orden"
        .Refresh
        .Recordset.Find "Orden=" & UltimoOrden
        If Not .Recordset.EOF Then
            .Recordset!Orden = .Recordset!Orden + RegIngresos + 1
            .Recordset.Update
        End If
                
        'Por fin, cambia el campo orden comenzando con el último registro que había de otros ingresos
        'es decir, con total 42. otros ingresos hasta llegar a Otros Ingresos
        'decrementa comenzando con ultimoorden + regingresos
        Decrementador = 1
        .Recordset.MoveFirst
        .Recordset.Find "Orden=" & UltReg
        Do While .Recordset!Orden <> PrimReg - 1
            .Recordset!Orden = UltimoOrden + RegIngresos + 1 - Decrementador
            .Recordset.Update
            .Recordset.MovePrevious
            Decrementador = Decrementador + 1
        Loop
        
        'borra 4. Ingresos y su Total
        .Recordset.MoveFirst
'        .Recordset.Find "descripcion like '4. ingresos'"
        .Recordset.Find "descripcion like '%" & MDIPrimero.AdoConfiguracion.Recordset!Ingresos & "%'"
        If Not .Recordset.EOF Then
            .Recordset.Delete
            .Recordset.Update
        End If
        .Recordset.MoveFirst
'        .Recordset.Find "descripcion like 'total 4. ingresos'"
        .Recordset.Find "descripcion like 'total " & MDIPrimero.AdoConfiguracion.Recordset!Ingresos & "'"
        If Not .Recordset.EOF Then
            .Recordset.Delete
            .Recordset.Update
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'aquí sería bueno mover los registros que estén entre costos y totales de ingresos operativos y mandarlos hasta abajito
'        Dim RegInicioCostosOperativos As Integer 'variable que me guarda el orden donde se encuentra el registro donde comienzan los costos operativos
'        Dim RegTotalCostosOperativos As Integer 'variable que me guarda el orden donde se encuentra el registro de total de costos operativos

        .Recordset.MoveFirst
        .Recordset.Find "descripcion like '" & MDIPrimero.AdoConfiguracion.Recordset!Costos & "'"
        If Not .Recordset.EOF Then
            RegInicioCostosOperativos = .Recordset!Orden
        End If
        .Recordset.MoveFirst
        .Recordset.Find "descripcion like '" & MDIPrimero.AdoConfiguracion.Recordset!Costos & "'"
        If Not .Recordset.EOF Then
            RegTotalCostosOperativos = .Recordset!Orden
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'comienzan los borrados de las agrupaciones de costos operativos y su total
        'borra 5. Costos y Gastos y su Total (las agrupaciones)
        .Recordset.MoveFirst
        .Recordset.Find "descripcion like '" & MDIPrimero.AdoConfiguracion.Recordset!Costos & "'"
                 
        If Not .Recordset.EOF Then
            .Recordset.Delete
            .Recordset.Update
        End If
        .Recordset.MoveFirst
        .Recordset.Find "descripcion like '%total " & MDIPrimero.AdoConfiguracion.Recordset!Costos & "%'"
        
        If Not .Recordset.EOF Then
            UltReg = .Recordset!Orden
            .Recordset.Delete
            .Recordset.Update
        End If
        
        'Utilidad Bruta
        .Recordset.MoveFirst
        '.Recordset.Find "descripcion like '51. costo de mercaderia'"
        .Recordset.Find "orden=" & PrimReg - 1 'encuentra total ingresos operativos
        Utilidad = (.Recordset!Haber1 - .Recordset!Debe1)
        Utilidad2 = (.Recordset!Haber2 - .Recordset!Debe2)
.Recordset.MoveFirst
        .Recordset.Find "descripcion like '%Total " & MDIPrimero.AdoConfiguracion.Recordset!CostosOperativos & "%'" 'encuentra total  de costos operativos
        Utilidad = Utilidad - (.Recordset!Haber1 - .Recordset!Debe1)
        Utilidad2 = Utilidad2 - (.Recordset!Haber2 - .Recordset!Debe2)
        RegTCostosOper = .Recordset!Orden
        .Recordset.MoveNext
        'encuentra el orden de total costos operativos donde será guardada la utilida bruta
        'empuja uno hacia adelante, es decir suma uno a orden para agregar la utilida bruta
        Do While Not .Recordset.EOF
            .Recordset!Orden = .Recordset!Orden + 1
            .Recordset.Update
            .Recordset.MoveNext
        Loop

        .Recordset.AddNew
            .Recordset!Descripcion = "UTILIDAD BRUTA"
            .Recordset!Orden = RegTCostosOper + 1
            .Recordset!Haber1 = Utilidad
            .Recordset!Haber2 = Utilidad2
            .Recordset!Haber3 = Utilidad + Utilidad2
            .Recordset!Nivel = 2
            .Recordset!KeyGrupo = "RP"
        .Recordset.Update
        .Recordset.MoveFirst
        .Recordset.Find "descripcion like '%total " & MDIPrimero.AdoConfiguracion.Recordset!Costos & "%'"
        If Not .Recordset.EOF Then
            .Recordset.Delete
            .Recordset.Update
        End If
End With
End If 'fin del si hay utilidad bruta
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'




End Sub
Public Sub SaldoReportesAcumulado(QUIEN As String)
Dim CodigoGrupo As String, Sql As String, Fechas As Date
Dim Nivel As Integer, Longitud As Integer, Fecha1 As String
Dim TotalMayor() As String, TotalDescripcion As String
Dim KeySuperior As String, NumeroHijos As Double, NumeroHijosTotales As Double
Dim DescripCuenta As String, DescripcionPadre As String, KeyUltimo As String, Ajuste As String
  '   ////////////////Elimino los registros del reporte///////////////////
  'frmreportes.DtaElimina.RecordSource = "DELETE Reportes.* From Reportes"
  'frmreportes.DtaElimina.Recordset.Updatable
  
    Dim Orden As Integer  'sirve para ordenar las cuentas
    Orden = 1
 
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 
         If FrmReportes.CmbMoneda.Text = "Córdobas" Then
            Ajuste = "Dólares"
         ElseIf FrmReportes.CmbMoneda.Text = "Dólares" Then
            Ajuste = "Córdobas"
         
         End If
         


'Busco que cuentas tienen saldo
 If QUIEN = "Balance" Then
  Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda " & vbLf
  Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Otros Activos' Or (Cuentas.TipoCuenta) = 'Caja' Or (Cuentas.TipoCuenta) = 'Bancos' Or (Cuentas.TipoCuenta) = 'Cuentas x Cobrar' Or (Cuentas.TipoCuenta) = 'Inventario' Or (Cuentas.TipoCuenta) = 'Papeleria - Utiles' Or (Cuentas.TipoCuenta) = 'Activo Fijo' Or (Cuentas.TipoCuenta) = 'Otros Pasivos' Or (Cuentas.TipoCuenta) = 'Cuentas x Pagar' Or (Cuentas.TipoCuenta) = 'Pasivo' Or (Cuentas.TipoCuenta) = 'Capital')) ORDER BY Cuentas.CodCuentas"
  FrmReportes.DtaHistorial.RecordSource = Sql
  FrmReportes.DtaHistorial.Refresh
 ElseIf QUIEN = "Utilidad" Then
  Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda " & vbLf
  Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Ingresos - Ventas' Or (Cuentas.TipoCuenta) = 'Costos' Or (Cuentas.TipoCuenta) = 'Gastos')) ORDER BY Cuentas.CodCuentas"
  FrmReportes.DtaHistorial.RecordSource = Sql
  FrmReportes.DtaHistorial.Refresh
 ElseIf QUIEN = "Resultado" Then
  Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda " & vbLf
  Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Ingresos - Ventas' Or (Cuentas.TipoCuenta) = 'Costos' Or (Cuentas.TipoCuenta) = 'Gastos')) ORDER BY Cuentas.CodCuentas"
  FrmReportes.DtaHistorial.RecordSource = Sql
  FrmReportes.DtaHistorial.Refresh
 ElseIf QUIEN = "UtilidadResultado" Then
  Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda " & vbLf
  Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Ingresos - Ventas' Or (Cuentas.TipoCuenta) = 'Costos' Or (Cuentas.TipoCuenta) = 'Gastos')) ORDER BY Cuentas.CodCuentas"
  FrmReportes.DtaHistorial.RecordSource = Sql
  FrmReportes.DtaHistorial.Refresh
 ElseIf QUIEN = "UtilidadAnterior" Then
'   SQL = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda " & vbLf
'  SQL = SQL & "Having (((Cuentas.TipoCuenta) = 'Ingresos - Ventas' Or (Cuentas.TipoCuenta) = 'Costos' Or (Cuentas.TipoCuenta) = 'Gastos')) ORDER BY Cuentas.CodCuentas"
  Sql = "SELECT  Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.Credito * Transacciones.TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta , Cuentas.TipoMoneda FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE  (Transacciones.FechaTransaccion < CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102)) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda HAVING (Cuentas.TipoCuenta = 'Ingresos - Ventas') OR (Cuentas.TipoCuenta = 'Costos') OR (Cuentas.TipoCuenta = 'Gastos') ORDER BY Cuentas.CodCuentas"
  FrmReportes.DtaHistorial.RecordSource = Sql
  FrmReportes.DtaHistorial.Refresh
'  QUIEN = "Utilidad"
 
 Else
  FrmReportes.DtaHistorial.RecordSource = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda ORDER BY Cuentas.CodCuentas"
'  InputBox "", "", FrmReportes.DtaHistorial.RecordSource
  FrmReportes.DtaHistorial.Refresh

 End If
' InputBox "", "", FrmReportes.DtaHistorial.RecordSource
 Totalingresos = 0
 TotalGastos = 0
 FrmReportes.osProgress1.Value = 0
 FrmReportes.osProgress1.Visible = True
 If Not FrmReportes.DtaHistorial.Recordset.EOF Then
  FrmReportes.DtaHistorial.Recordset.MoveLast
  FrmReportes.osProgress1.Max = FrmReportes.DtaHistorial.Recordset.RecordCount
  FrmReportes.DtaHistorial.Recordset.MoveFirst
 End If
  Do While Not FrmReportes.DtaHistorial.Recordset.EOF
  
   NumFecha1 = FechaIni
   NumFecha2 = FechaFin
   
   If QUIEN = "UtilidadAnterior" Then
        FechaIni = "01/01/" & Year(FechaIni) - 1
        FechaFin = "31/12/" & Year(FechaFin) - 1
   
        NumFecha1 = FechaIni
        NumFecha2 = FechaFin
 
   
   End If
   
       FrmReportes.LblProgreso.Caption = "Consultando Registros del periodo seleccionado para la cuenta " & CodigoCuenta

          FrmReportes.osProgress1.Value = FrmReportes.osProgress1.Value + 1


    '///////////////////////////////////////////////////////////////////////////
    '///////////////////////REGISTRO DEL PERIODO SELECCIONADO./////////////////
    '///////////////////////////////////////////////////////////////////////////
    
    CodigoCuenta = FrmReportes.DtaHistorial.Recordset("CodCuentas")
    
    If CodigoCuenta = "6-51 " Then
      CodigoCuenta = "6-51 "
    End If
'    FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Transacciones.FechaTransaccion, Tasas.MontoCordobas, Tasas.MontoLibras, Transacciones.NTransaccion FROM Tasas INNER JOIN (Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas) ON Tasas.FechaTasas = Transacciones.FechaTasas GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Transacciones.FechaTransaccion, Tasas.MontoCordobas, Tasas.MontoLibras, Transacciones.NTransaccion HAVING (((Cuentas.CodCuentas)='" & CodigoCuenta & "') AND ((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "')) ORDER BY Cuentas.CodCuentas"
    FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.Credito * Transacciones.TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta , Cuentas.TipoMoneda, Transacciones.FechaTransaccion, Tasas.MontoCordobas, Tasas.MontoLibras, Transacciones.NTransaccion FROM  Tasas INNER JOIN Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Tasas.FechaTasas = Transacciones.FechaTasas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento And Transacciones.NPeriodo = IndiceTransaccion.NPeriodo GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Transacciones.FechaTransaccion, Tasas.MontoCordobas, " & _
                                           "Tasas.MontoLibras , Transacciones.NTransaccion, IndiceTransaccion.Ajuste HAVING (Cuentas.CodCuentas = '" & CodigoCuenta & "') AND (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') AND (IndiceTransaccion.Ajuste <> '" & Ajuste & "') ORDER BY Cuentas.CodCuentas"
    FrmReportes.DtaConsulta.Refresh
    
    
    DoEvents
    
    TotalCuenta = 0
    Total1 = 0
    FrmReportes.osProgress2.Value = 0
    Do While Not FrmReportes.DtaConsulta.Recordset.EOF
      FrmReportes.osProgress2.Visible = True
        If FrmReportes.osProgress2.Value = 0 Then
'            FrmReportes.lblProgreso.Caption = "Consultando Registros del periodo seleccionado para la cuenta " & CodigoCuenta
            FrmReportes.osProgress2.Max = FrmReportes.DtaConsulta.Recordset.RecordCount
            FrmReportes.osProgress2.Value = FrmReportes.osProgress2.Value + 1

             DoEvents
        Else
'            FrmReportes.lblProgreso.Caption = "Consultando Registros del periodo seleccionado para la cuenta " & CodigoCuenta
            FrmReportes.osProgress2.Value = FrmReportes.osProgress2.Value + 1
            DoEvents
        End If

        
       TotalDebito = 0
       TotalCredito = 0
      TipoCuenta = FrmReportes.DtaConsulta.Recordset("TipoCuenta")
      TipoMoneda = FrmReportes.DtaConsulta.Recordset("TipoMoneda")
      FechaTransaccion = FrmReportes.DtaConsulta.Recordset("FechaTransaccion")
      NumFecha = FrmReportes.DtaConsulta.Recordset("FechaTransaccion")
      Fechas = FrmReportes.DtaConsulta.Recordset("FechaTransaccion")
      FrmReportes.DtaTasas2.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas WHERE (FechaTasas = '" & Format(Fechas, "yyyymmdd") & "')"
      FrmReportes.DtaTasas2.Refresh
      TasaCambio = FrmReportes.DtaTasas2.Recordset("MontoCordobas")
     If TasaCambio = 0 Then
      cadena = "La tasa de Cambio con Fecha: " & Fechas1 & vbLf
      cadena = cadena & "no puede ser igual a Cero, el Sistema Contable" & vbLf
      cadena = cadena & "no contiuara el proceso......"
      MsgBox cadena, vbCritical, "Sistema Contable"
      Exit Sub
     End If
      If Not IsNull(FrmReportes.DtaConsulta.Recordset("DescripcionCuentas")) Then
       Descripcion = FrmReportes.DtaConsulta.Recordset("CodCuentas") + "." + FrmReportes.DtaConsulta.Recordset("DescripcionCuentas")
      Else
       Descripcion = FrmReportes.DtaConsulta.Recordset("CodCuentas") + "." + "NO TIENE DESCRIPCION????"
      End If
      If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
        If Not IsNull(FrmReportes.DtaConsulta.Recordset("MDebito")) Then
            Debito = FrmReportes.DtaConsulta.Recordset("MDebito")
        End If
        If Not IsNull(FrmReportes.DtaConsulta.Recordset("MCredito")) Then
            Credito = FrmReportes.DtaConsulta.Recordset("MCredito")
        End If
        Total1 = Debito - Credito + Total1
        
        '//////////////////Verifico el tipo de moneda de la cuenta////////////////////
        If Not TipoMoneda = FrmReportes.CmbMoneda.Text Then
           Select Case TipoMoneda
              Case "Córdobas"
                    TotalCuenta = (Debito - Credito) / TasaCambio + TotalCuenta
          
              Case "Dólares"
                    TotalCuenta = (Debito - Credito) * TasaCambio + TotalCuenta
           
          End Select
        Else
               TotalCuenta = (Debito - Credito) + TotalCuenta
        End If

        Debito = 0
        Credito = 0
      Else
         If Not IsNull(FrmReportes.DtaConsulta.Recordset("MDebito")) Then
            Debito = FrmReportes.DtaConsulta.Recordset("MDebito")
         End If
         If Not IsNull(FrmReportes.DtaConsulta.Recordset("MCredito")) Then
            Credito = FrmReportes.DtaConsulta.Recordset("MCredito")
         End If
         
        '//////////////////Verifico el tipo de moneda de la cuenta////////////////////
        If Not TipoMoneda = FrmReportes.CmbMoneda.Text Then
           Select Case TipoMoneda
              Case "Córdobas"
                    TotalCuenta = (Credito - Debito) / TasaCambio + TotalCuenta
          
              Case "Dólares"
                    TotalCuenta = (Credito - Debito) * TasaCambio + TotalCuenta
           
          End Select
        Else
               TotalCuenta = (Credito - Debito) + TotalCuenta
        End If
         
         Total1 = Credito - Debito + Total1
         Debito = 0
         Credito = 0
      End If
    
    

 


   
   FrmReportes.DtaConsulta.Recordset.MoveNext

   Loop

'/////////////////////////////////////////////////////////////////////////////////////
'////////////////GRABO LOS REGISTROS DEL PERIDODO SELECCIONADO///////////////////////
'//////////////////////////////////////////////////////////////////////////////////////
        
    If QUIEN = "Balanza" Then
'//////////////////////////Busco la Cuenta en Reportes/////////////////////////////////
       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.KeyGrupoCuenta,Reportes.KeyGrupoSuperior,Reportes.KeyGrupo,Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3 From Reportes Where (((Reportes.KeyGrupo) = '" & CodigoCuenta & "'))"
       FrmReportes.DtaConsulta2.Refresh
       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
'guillermo, total debe y total haber
'          'FrmReportes.DtaConsulta2.Recordset.Edit
          If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
           'guillermo, total debe y total haber
           If TotalCuenta < 0 Then
             FrmReportes.DtaConsulta2.Recordset("Haber2") = Abs(TotalCuenta)
           Else
             FrmReportes.DtaConsulta2.Recordset("Debe2") = TotalCuenta
           End If
          Else
           If TotalCuenta < 0 Then
               FrmReportes.DtaConsulta2.Recordset("Debe2") = Abs(TotalCuenta)
           Else
               FrmReportes.DtaConsulta2.Recordset("Haber2") = TotalCuenta
           End If
          End If
          FrmReportes.DtaConsulta2.Recordset.Update
       End If
       
'//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
              Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta"))
              If Longitud > 1 Then
                 If Longitud = 5 Then
                     Nivel = 2
                 Else
                     Nivel = (Longitud - 5) / 2
                     Nivel = Nivel + 2
                 End If
              Else
                    Nivel = 1
              End If
       
'///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
       Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta")
       'Nivel = Nivel - 1
       For i = Nivel To 1 Step -1
'/////////Busco el Grupo para Sumar los Totaldes
           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = '" & Respuesta & "'))"
           FrmReportes.DtaConsulta.Refresh
           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
              FrmReportes.DtaConsulta.Recordset.MoveLast
'              'FrmReportes.'DtaConsulta.Recordset.Edit
              'guillermo, total debe y total haber, son los superiores
              If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                If (TotalCuenta) < 0 Then
                  FrmReportes.DtaConsulta.Recordset("Haber2") = FrmReportes.DtaConsulta.Recordset("Haber2") + Abs(TotalCuenta)
                Else
                 FrmReportes.DtaConsulta.Recordset("Debe2") = FrmReportes.DtaConsulta.Recordset("Debe2") + TotalCuenta
                End If
              Else
               If TotalCuenta < 0 Then
                 FrmReportes.DtaConsulta.Recordset("Debe2") = FrmReportes.DtaConsulta.Recordset("Debe2") + Abs(TotalCuenta)
               Else
                 FrmReportes.DtaConsulta.Recordset("Haber2") = FrmReportes.DtaConsulta.Recordset("Haber2") + TotalCuenta
               End If
              End If
              FrmReportes.DtaConsulta.Recordset.Update

           End If

           If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
              Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
           End If
       Next
    
    
    
    ElseIf QUIEN = "BalanzaCodigo" Then
    
     '////////////Agrego los Saldos del Periodo Seleccionado////////////////////
'      FrmReportes.lblProgreso.Caption = "Agregando saldos al Periodo"
'      FrmReportes.osProgress1.Value = 0
'      FrmReportes.osProgress1.Max
'
      FrmReportes.DtaReportes.Recordset.AddNew
      FrmReportes.DtaReportes.Recordset("Descripcion") = Descripcion
      If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
        If TotalCuenta < 0 Then
         FrmReportes.DtaReportes.Recordset("Haber2") = Abs(TotalCuenta)
        ElseIf TotalCuenta > 0 Then
         FrmReportes.DtaReportes.Recordset("Debe2") = TotalCuenta
        End If
         GranTotal = GranTotal + TotalCuenta
      Else
        If TotalCuenta < 0 Then
          FrmReportes.DtaReportes.Recordset("Debe2") = Abs(TotalCuenta)
        ElseIf TotalCuenta > 0 Then
          FrmReportes.DtaReportes.Recordset("Haber2") = TotalCuenta
        End If
         GranTotal = GranTotal + TotalCuenta
      End If
      FrmReportes.DtaReportes.Recordset!Orden = Orden
      Orden = Orden + 1
      FrmReportes.DtaReportes.Recordset.Update
      
    
    ElseIf QUIEN = "Balance" Then
    
'//////////////////////////Busco la Cuenta en Reportes/////////////////////////////////
       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.KeyGrupoCuenta,Reportes.KeyGrupoSuperior,Reportes.KeyGrupo,Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3 From Reportes Where (((Reportes.KeyGrupo) = '" & CodigoCuenta & "'))ORDER BY Orden "
       FrmReportes.DtaConsulta2.Refresh
       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then

          'FrmReportes.DtaConsulta2.Recordset.Edit
            FrmReportes.DtaConsulta2.Recordset("Debe2") = TotalCuenta
          FrmReportes.DtaConsulta2.Recordset.Update
       End If
       
'//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
              Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta"))
              If Longitud > 1 Then
                 If Longitud = 5 Then
                     Nivel = 2
                 Else
                     Nivel = (Longitud - 5) / 2
                     Nivel = Nivel + 2
                 End If
              Else
                    Nivel = 1
              End If
       
'///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
       Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta")
       'Nivel = Nivel - 1
       For i = Nivel To 1 Step -1
'/////////Busco el Grupo para Sumar los Totaldes
           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = '" & Respuesta & "')) ORDER BY Orden"
'          InputBox "", "", FrmReportes.DtaConsulta.RecordSource
           FrmReportes.DtaConsulta.Refresh
           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
              FrmReportes.DtaConsulta.Recordset.MoveLast
'              'FrmReportes.'DtaConsulta.Recordset.Edit
                 FrmReportes.DtaConsulta.Recordset("Haber2") = FrmReportes.DtaConsulta.Recordset("Haber2") + TotalCuenta
              FrmReportes.DtaConsulta.Recordset.Update

           End If

           If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
              Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
           End If
       Next
'//UTILIDAD
    ElseIf QUIEN = "UtilidadResultado" Then
       If TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Then
         TotalGastos = TotalGastos + TotalCuenta
    
       ElseIf TipoCuenta = "Ingresos - Ventas" Then
         Totalingresos = Totalingresos + TotalCuenta
       End If
  '       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta From Reportes Where (((Reportes.Descripcion) Like '*Resultado Periodo*'))"
       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.Descripcion) Like '%Resultado Periodo%'))"
       FrmReportes.DtaConsulta2.Refresh
       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
           'FrmReportes.DtaConsulta2.Recordset.Edit
'               FrmReportes.DtaConsulta2.Recordset("Haber1") = Totalingresos - TotalGastos
               FrmReportes.DtaConsulta2.Recordset("Haber2") = Totalingresos - TotalGastos
           FrmReportes.DtaConsulta2.Recordset.Update
       End If
      
    
    ElseIf QUIEN = "Utilidad" Then
       If TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Then
         TotalGastos = TotalGastos + TotalCuenta
    
       ElseIf TipoCuenta = "Ingresos - Ventas" Then
         Totalingresos = Totalingresos + TotalCuenta
       End If

       FrmReportes.DtaConsulta2.RecordSource = "SELECT Descripcion, Debe1, Haber1, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Haber2, Debe2, Debe3, Haber3, Nivel, Orden From Reportes WHERE     (Descripcion LIKE '%Resultado Periodo%')"
'       InputBox "", "", FrmReportes.DtaConsulta2.RecordSource
       FrmReportes.DtaConsulta2.Refresh
       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
           'FrmReportes.DtaConsulta2.Recordset.Edit
               FrmReportes.DtaConsulta2.Recordset("Debe2") = Totalingresos - TotalGastos
           FrmReportes.DtaConsulta2.Recordset.Update
    
'//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
              Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupo"))
              If Longitud > 1 Then
                 If Longitud = 5 Then
                     Nivel = 2
                 Else
                     Nivel = (Longitud - 5) / 2
                     Nivel = Nivel + 2
                 End If
              Else
                    Nivel = 1
              End If
       
'///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
       Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupo")
       'Nivel = Nivel - 1
'       For I = Nivel To 1 Step -1
'/////////Busco el Grupo para Sumar los Totaldes
           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = 'PC')) ORDER BY Orden"
           FrmReportes.DtaConsulta.Refresh
           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
              FrmReportes.DtaConsulta.Recordset.MoveLast
                 FrmReportes.DtaConsulta.Recordset("Haber2") = (Totalingresos - TotalGastos)
              FrmReportes.DtaConsulta.Recordset.Update

           End If
           
           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = 'C')) ORDER BY Orden"
           FrmReportes.DtaConsulta.Refresh
           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
              FrmReportes.DtaConsulta.Recordset.MoveLast
                 FrmReportes.DtaConsulta.Recordset("Haber2") = (Totalingresos - TotalGastos)
              FrmReportes.DtaConsulta.Recordset.Update

           End If

'           If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
'              Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
'           End If
'       Next
      End If
    ElseIf QUIEN = "Resultado" Then
'//////////////////////////Busco la Cuenta en Reportes/////////////////////////////////
       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.KeyGrupoCuenta,Reportes.KeyGrupoSuperior,Reportes.KeyGrupo,Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3 From Reportes Where (((Reportes.KeyGrupo) = '" & CodigoCuenta & "'))"
       FrmReportes.DtaConsulta2.Refresh
       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then

          'FrmReportes.DtaConsulta2.Recordset.Edit
            FrmReportes.DtaConsulta2.Recordset("Debe2") = TotalCuenta
          FrmReportes.DtaConsulta2.Recordset.Update
'       End If '////pruekdkdk
       
'//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
              Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta"))
              If Longitud > 1 Then
                 If Longitud = 5 Then
                     Nivel = 2
                 Else
                     Nivel = (Longitud - 5) / 2
                     Nivel = Nivel + 2
                 End If
              Else
                    Nivel = 1
              End If
       
'///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
       Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta")
       Else
        MsgBox "La cuenta Tiene Saldo y no aparece en la Estructura, Cuenta: " & CodigoCuenta, vbCritical
       End If
       'Nivel = Nivel - 1
       For i = Nivel To 1 Step -1
'/////////Busco el Grupo para Sumar los Totaldes
           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = '" & Respuesta & "')) ORDER BY Orden"
'           InputBox "", "", DtaConsulta.RecordSource
           FrmReportes.DtaConsulta.Refresh
           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
              FrmReportes.DtaConsulta.Recordset.MoveLast
              'FrmReportes.'DtaConsulta.Recordset.Edit
                 FrmReportes.DtaConsulta.Recordset("Haber2") = FrmReportes.DtaConsulta.Recordset("Haber2") + TotalCuenta
              FrmReportes.DtaConsulta.Recordset.Update

           End If

           If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
              Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
           End If
       Next
    End If
    

   
   FrmReportes.DtaHistorial.Recordset.MoveNext
  Loop
  
  FrmReportes.osProgress2.Visible = False
 '/////////////////////////////////////////////////////////////////////////////////////////////////////
  '//////////////////PERIODO ANTERIOR///////////////////////////////////////////////////////////////
  '////////////////////////////////////////////////////////////////////////////////////////////////////
  
  Debito = 0
  Credito = 0
  Totalingresos = 0
  TotalGastos = 0
  'Busco que cuentas tienen saldo
 If QUIEN = "Utilidad" Then
   Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) <'" & Format(FechaIni, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda" & vbLf
   Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Ingresos - Ventas' Or (Cuentas.TipoCuenta) = 'Costos' Or (Cuentas.TipoCuenta) = 'Gastos')) ORDER BY Cuentas.CodCuentas"
 ElseIf QUIEN = "Resultado" Then
   Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) <'" & Format(FechaIni, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda" & vbLf
   Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Ingresos - Ventas' Or (Cuentas.TipoCuenta) = 'Costos' Or (Cuentas.TipoCuenta) = 'Gastos')) ORDER BY Cuentas.CodCuentas"
 ElseIf QUIEN = "UtilidadResultado" Then
  Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) <'" & Format(FechaIni, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda " & vbLf
  Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Ingresos - Ventas' Or (Cuentas.TipoCuenta) = 'Costos' Or (Cuentas.TipoCuenta) = 'Gastos')) ORDER BY Cuentas.CodCuentas"
  FrmReportes.DtaHistorial.RecordSource = Sql
  FrmReportes.DtaHistorial.Refresh
 Else
  Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) <'" & Format(FechaIni, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda" & vbLf
  Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Otros Activos' Or (Cuentas.TipoCuenta) = 'Caja' Or (Cuentas.TipoCuenta) = 'Bancos' Or (Cuentas.TipoCuenta) = 'Cuentas x Cobrar' Or (Cuentas.TipoCuenta) = 'Inventario' Or (Cuentas.TipoCuenta) = 'Papeleria - Utiles' Or (Cuentas.TipoCuenta) = 'Activo Fijo' Or (Cuentas.TipoCuenta) = 'Otros Pasivos' Or (Cuentas.TipoCuenta) = 'Cuentas x Pagar' Or (Cuentas.TipoCuenta) = 'Pasivo' Or (Cuentas.TipoCuenta) = 'Capital')) ORDER BY Cuentas.CodCuentas"
 End If
 FrmReportes.DtaHistorial.RecordSource = Sql

'InputBox "", "", FrmReportes.DtaHistorial.RecordSource
 FrmReportes.DtaHistorial.Refresh

FrmReportes.LblProgreso.Caption = "Consultando Registros del Periodo Anterior para " & QUIEN
FrmReportes.osProgress1.Value = 0

 If Not FrmReportes.DtaHistorial.Recordset.EOF Then
  FrmReportes.osProgress1.Max = FrmReportes.DtaHistorial.Recordset.RecordCount
  FrmReportes.DtaHistorial.Refresh
 Else
'  Exit Sub
 End If
 Do While Not FrmReportes.DtaHistorial.Recordset.EOF
    '////////Consulto los registros del periodo ANTERIOR.///////////
    FrmReportes.osProgress1.Value = FrmReportes.osProgress1.Value + 1
    CodigoCuenta = FrmReportes.DtaHistorial.Recordset("CodCuentas")
'    FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Transacciones.FechaTransaccion, Tasas.MontoCordobas, Tasas.MontoLibras, Transacciones.NTransaccion FROM Tasas INNER JOIN (Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas) ON Tasas.FechaTasas = Transacciones.FechaTasas GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Transacciones.FechaTransaccion, Tasas.MontoCordobas, Tasas.MontoLibras, Transacciones.NTransaccion HAVING (((Cuentas.CodCuentas)='" & CodigoCuenta & "') AND ((Transacciones.FechaTransaccion) <'" & Format(FechaIni, "yyyymmdd") & "')) ORDER BY Cuentas.CodCuentas"
    FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.Credito * Transacciones.TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta , Cuentas.TipoMoneda, Transacciones.FechaTransaccion, Tasas.MontoCordobas, Tasas.MontoLibras, Transacciones.NTransaccion FROM Tasas INNER JOIN Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Tasas.FechaTasas = Transacciones.FechaTasas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND  Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento And Transacciones.NPeriodo = IndiceTransaccion.NPeriodo GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Transacciones.FechaTransaccion, Tasas.MontoCordobas, " & _
                                           "Tasas.MontoLibras , Transacciones.NTransaccion, IndiceTransaccion.Ajuste HAVING  (Cuentas.CodCuentas = '" & CodigoCuenta & "') AND (Transacciones.FechaTransaccion < '" & Format(FechaIni, "yyyymmdd") & "') AND (IndiceTransaccion.Ajuste <> '" & Ajuste & "') ORDER BY Cuentas.CodCuentas"
    FrmReportes.DtaConsulta.Refresh
    TotalCuenta = 0
    Total1 = 0
    Do While Not FrmReportes.DtaConsulta.Recordset.EOF

       TotalDebito = 0
       TotalCredito = 0
      TipoCuenta = FrmReportes.DtaConsulta.Recordset("TipoCuenta")
      TipoMoneda = FrmReportes.DtaConsulta.Recordset("TipoMoneda")
      FechaTransaccion = FrmReportes.DtaConsulta.Recordset("FechaTransaccion")
      Fechas1 = FrmReportes.DtaConsulta.Recordset("FechaTransaccion")
      TasaCambio = FrmReportes.DtaConsulta.Recordset("MontoCordobas")
     If TasaCambio = 0 Then
      cadena = "La tasa de Cambio de Cambio con Fecha: " & Fechas1 & vbLf
      cadena = cadena & "no puede ser igual a Cero, el Sistema Contable" & vbLf
      cadena = cadena & "no contiuara el proceso......"
      MsgBox cadena, vbCritical, "Sistema Contable"
      Exit Sub
     End If
      
        If Not IsNull(FrmReportes.DtaConsulta.Recordset("DescripcionCuentas")) Then
         Descripcion = FrmReportes.DtaConsulta.Recordset("CodCuentas") + "." + FrmReportes.DtaConsulta.Recordset("DescripcionCuentas")
        Else
         Descripcion = FrmReportes.DtaConsulta.Recordset("CodCuentas")
        End If
      If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
        If Not IsNull(FrmReportes.DtaConsulta.Recordset("MDebito")) Then
            Debito = FrmReportes.DtaConsulta.Recordset("MDebito")
        End If
        If Not IsNull(FrmReportes.DtaConsulta.Recordset("MCredito")) Then
            Credito = FrmReportes.DtaConsulta.Recordset("MCredito")
        End If
        Total1 = Debito - Credito + Total1
        
        '//////////////////Verifico el tipo de moneda de la cuenta////////////////////
        If Not TipoMoneda = FrmReportes.CmbMoneda.Text Then
           Select Case TipoMoneda
              Case "Córdobas"
                    TotalCuenta = (Debito - Credito) / TasaCambio + TotalCuenta
          
              Case "Dólares"
                    TotalCuenta = (Debito - Credito) * TasaCambio + TotalCuenta
           
          End Select
        Else
               TotalCuenta = (Debito - Credito) + TotalCuenta
        End If

        Debito = 0
        Credito = 0
      Else
         If Not IsNull(FrmReportes.DtaConsulta.Recordset("MDebito")) Then
            Debito = FrmReportes.DtaConsulta.Recordset("MDebito")
         End If
         If Not IsNull(FrmReportes.DtaConsulta.Recordset("MCredito")) Then
            Credito = FrmReportes.DtaConsulta.Recordset("MCredito")
         End If
         
        '//////////////////Verifico el tipo de moneda de la cuenta////////////////////
        If Not TipoMoneda = FrmReportes.CmbMoneda.Text Then
           Select Case TipoMoneda
              Case "Córdobas"
                    TotalCuenta = (Credito - Debito) / TasaCambio + TotalCuenta
          
              Case "Dólares"
                    TotalCuenta = (Credito - Debito) * TasaCambio + TotalCuenta
           
          End Select
        Else
               TotalCuenta = (Credito - Debito) + TotalCuenta
        End If
         
         Total1 = Credito - Debito + Total1
         Debito = 0
         Credito = 0
      End If
    
    

 


   
   FrmReportes.DtaConsulta.Recordset.MoveNext
   
   Loop
   
 '//////////////////////////////////////////////////////////////////////
 '////////////////////GRABO LOS REGISTROS DEL PERIODO ANTERIOR//////////
 '//////////////////EN LA TABLA REPORTES////////////////////////////////
 
    If QUIEN = "Balanza" Then
'//////////////////////////Busco la Cuenta en Reportes/////////////////////////////////
       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.KeyGrupoCuenta,Reportes.KeyGrupoSuperior,Reportes.KeyGrupo,Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3,reportes.orden From Reportes Where (((Reportes.KeyGrupo) = '" & CodigoCuenta & "'))"
       FrmReportes.DtaConsulta2.Refresh
       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then

          'FrmReportes.DtaConsulta2.Recordset.Edit
          If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
            If TotalCuenta < 0 Then
             FrmReportes.DtaConsulta2.Recordset("Haber1") = Abs(TotalCuenta)
            Else
             FrmReportes.DtaConsulta2.Recordset("Debe1") = TotalCuenta
            End If
          Else
            If TotalCuenta < 0 Then
              FrmReportes.DtaConsulta2.Recordset("Debe1") = Abs(TotalCuenta)
            Else
             FrmReportes.DtaConsulta2.Recordset("Haber1") = TotalCuenta
            End If
          End If
'          FrmReportes.DtaConsulta2.Recordset!Orden = Orden
          Orden = Orden + 1
          FrmReportes.DtaConsulta2.Recordset.Update
       End If
       
'//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
              Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta"))
              If Longitud > 1 Then
                 If Longitud = 5 Then
                     Nivel = 2
                 Else
                     Nivel = (Longitud - 5) / 2
                     Nivel = Nivel + 2
                 End If
              Else
                    Nivel = 1
              End If
       
'///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
       Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta")
       'Nivel = Nivel - 1
       For i = Nivel To 1 Step -1
'/////////Busco el Grupo para Sumar los Totaldes
           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior,reportes.orden From Reportes Where (((Reportes.KeyGrupo) = '" & Respuesta & "'))"
           FrmReportes.DtaConsulta.Refresh
           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
              FrmReportes.DtaConsulta.Recordset.MoveLast
              'FrmReportes.'DtaConsulta.Recordset.Edit
              If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                If TotalCuenta < 0 Then
                  FrmReportes.DtaConsulta.Recordset("Haber1") = FrmReportes.DtaConsulta.Recordset("Haber1") + Abs(TotalCuenta)
                Else
                 FrmReportes.DtaConsulta.Recordset("Debe1") = FrmReportes.DtaConsulta.Recordset("Debe1") + TotalCuenta
                End If
              Else
                If TotalCuenta < 0 Then
                 FrmReportes.DtaConsulta.Recordset("Debe1") = FrmReportes.DtaConsulta.Recordset("Debe1") + Abs(TotalCuenta)
                Else
                 FrmReportes.DtaConsulta.Recordset("Haber1") = FrmReportes.DtaConsulta.Recordset("Haber1") + TotalCuenta
                End If
              End If
              If 1 = 0 Then FrmReportes.DtaConsulta.Recordset!Orden = Orden + 1
              Orden = Orden + 1
              
              FrmReportes.DtaConsulta.Recordset.Update

           End If

           If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
              Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
           End If
       Next
    
    
    
    ElseIf QUIEN = "BalanzaCodigo" Then
    
      '////////////Agrego los Saldos del PeriodoAnterior////////////////////
      
       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3,reportes.orden From Reportes Where (((Reportes.Descripcion) = '" & Descripcion & "'))"
       FrmReportes.DtaConsulta2.Refresh
       If FrmReportes.DtaConsulta2.Recordset.EOF Then
         FrmReportes.DtaConsulta2.Recordset.AddNew
         FrmReportes.DtaConsulta2.Recordset("Descripcion") = Descripcion
       Else
         'FrmReportes.DtaConsulta2.Recordset.Edit
       End If
      
       If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
         If TotalCuenta < 0 Then
           FrmReportes.DtaConsulta2.Recordset("Haber1") = Abs(TotalCuenta)
         Else
          FrmReportes.DtaConsulta2.Recordset("Debe1") = TotalCuenta
         End If
       Else
         If TotalCuenta < 0 Then
          FrmReportes.DtaConsulta2.Recordset("Debe1") = Abs(TotalCuenta)
         Else
          FrmReportes.DtaConsulta2.Recordset("Haber1") = TotalCuenta
         End If
       End If
       FrmReportes.DtaConsulta2.Recordset!Orden = Orden
       Orden = Orden + 1
       FrmReportes.DtaConsulta2.Recordset.Update
       
   
   '////////////////////AGREGO LOS SALDOS ANTERIORES DEL BALANCE///////////
   ElseIf QUIEN = "Balance" Then
'//////////////////////////Busco la Cuenta en Reportes/////////////////////////////////
       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.KeyGrupoCuenta,Reportes.KeyGrupoSuperior,Reportes.KeyGrupo,Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3 From Reportes Where (((Reportes.KeyGrupo) = '" & CodigoCuenta & "'))ORDER BY Orden "
       FrmReportes.DtaConsulta2.Refresh
       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then

          'FrmReportes.DtaConsulta2.Recordset.Edit
            FrmReportes.DtaConsulta2.Recordset("Debe1") = TotalCuenta + FrmReportes.DtaConsulta2.Recordset("Debe1")
          FrmReportes.DtaConsulta2.Recordset.Update
       End If

'//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
              Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta"))
              If Longitud > 1 Then
                 If Longitud = 5 Then
                     Nivel = 2
                 Else
                     Nivel = (Longitud - 5) / 2
                     Nivel = Nivel + 2
                 End If
              Else
                    Nivel = 1
              End If

'///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
       Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta")
       'Nivel = Nivel - 1
       For i = Nivel To 1 Step -1
'/////////Busco el Grupo para Sumar los Totaldes
           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = '" & Respuesta & "')) ORDER BY Orden"
'          InputBox "", "", FrmReportes.DtaConsulta.RecordSource
           FrmReportes.DtaConsulta.Refresh
           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
              FrmReportes.DtaConsulta.Recordset.MoveLast
'              'FrmReportes.'DtaConsulta.Recordset.Edit
                 FrmReportes.DtaConsulta.Recordset("Haber1") = FrmReportes.DtaConsulta.Recordset("Haber1") + TotalCuenta
              FrmReportes.DtaConsulta.Recordset.Update

           End If

           If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
              Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
           End If
       Next
   
 
 '//////////////////AGREGO LA UTILIDAD AL BALANCE//////////////////////////
 ElseIf QUIEN = "Utilidad" Then
       If TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Then
         TotalGastos = TotalCuenta
         Totalingresos = 0
       ElseIf TipoCuenta = "Ingresos - Ventas" Then
         Totalingresos = TotalCuenta
         TotalGastos = 0
       End If
       
       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta From Reportes Where (((Reportes.Descripcion) Like '%Resultado Periodo%'))"
       FrmReportes.DtaConsulta2.Refresh
       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
               FrmReportes.DtaConsulta2.Recordset("Debe1") = Totalingresos - TotalGastos + FrmReportes.DtaConsulta2.Recordset("Debe1")
           FrmReportes.DtaConsulta2.Recordset.Update
    
'//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
              Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupo"))
              If Longitud > 1 Then
                 If Longitud = 5 Then
                     Nivel = 2
                 Else
                     Nivel = (Longitud - 5) / 2
                     Nivel = Nivel + 2
                 End If
              Else
                    Nivel = 1
              End If
       
'///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
       Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupo")
       'Nivel = Nivel - 1
'       For I = Nivel To 1 Step -1
'/////////Busco el Grupo para Sumar los Totaldes
           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = 'PC')) ORDER BY Orden"
           FrmReportes.DtaConsulta.Refresh
           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
              FrmReportes.DtaConsulta.Recordset.MoveLast
                 FrmReportes.DtaConsulta.Recordset("Haber1") = (Totalingresos - TotalGastos) + FrmReportes.DtaConsulta.Recordset("Haber1")
              FrmReportes.DtaConsulta.Recordset.Update

           End If
           
           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = 'C')) ORDER BY Orden"
           FrmReportes.DtaConsulta.Refresh
           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
              FrmReportes.DtaConsulta.Recordset.MoveLast
                 FrmReportes.DtaConsulta.Recordset("Haber1") = (Totalingresos - TotalGastos) + FrmReportes.DtaConsulta.Recordset("Haber1")
              FrmReportes.DtaConsulta.Recordset.Update

           End If

'           If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
'              Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
'           End If
'       Next
      End If
  
ElseIf QUIEN = "UtilidadResultado" Then
       If TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Then
         TotalGastos = TotalGastos + TotalCuenta
    
       ElseIf TipoCuenta = "Ingresos - Ventas" Then
         Totalingresos = Totalingresos + TotalCuenta
       End If
  '       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta From Reportes Where (((Reportes.Descripcion) Like '*Resultado Periodo*'))"
       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.Descripcion) Like '%Resultado Periodo%'))"
       FrmReportes.DtaConsulta2.Refresh
       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
           'FrmReportes.DtaConsulta2.Recordset.Edit
               FrmReportes.DtaConsulta2.Recordset("Haber1") = Totalingresos - TotalGastos
           FrmReportes.DtaConsulta2.Recordset.Update
       End If
 
 
 ElseIf QUIEN = "Resultado" Then

'//////////////////////////Busco la Cuenta en Reportes/////////////////////////////////
       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.KeyGrupoCuenta,Reportes.KeyGrupoSuperior,Reportes.KeyGrupo,Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3 From Reportes Where (((Reportes.KeyGrupo) = '" & CodigoCuenta & "'))"
       FrmReportes.DtaConsulta2.Refresh
       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then

          'FrmReportes.DtaConsulta2.Recordset.Edit
            FrmReportes.DtaConsulta2.Recordset("Debe1") = TotalCuenta
          FrmReportes.DtaConsulta2.Recordset.Update
'       End If  '/////////DA UN ERROR SE INCLUYE KEYGRUPO
       
'//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
              Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta"))
              If Longitud > 1 Then
                 If Longitud = 5 Then
                     Nivel = 2
                 Else
                     Nivel = (Longitud - 5) / 2
                     Nivel = Nivel + 2
                 End If
              Else
                    Nivel = 1
              End If
       
'///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
       Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta")
       End If
       'Nivel = Nivel - 1
       For i = Nivel To 1 Step -1
'/////////Busco el Grupo para Sumar los Totaldes
           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = '" & Respuesta & "')) ORDER BY Orden"
'           InputBox "", "", DtaConsulta.RecordSource
           FrmReportes.DtaConsulta.Refresh
           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
              FrmReportes.DtaConsulta.Recordset.MoveLast
              'FrmReportes.'DtaConsulta.Recordset.Edit
                 FrmReportes.DtaConsulta.Recordset("Haber1") = FrmReportes.DtaConsulta.Recordset("Haber1") + TotalCuenta
              FrmReportes.DtaConsulta.Recordset.Update

           End If

           If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
              Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
           End If
       Next


 End If

   
   FrmReportes.DtaHistorial.Recordset.MoveNext
  Loop
  
'/////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////SUMO LOS TOTALES DEL PASIVO + CAPITAL/////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////
  
If QUIEN = "Balance" Then
   FrmReportes.DtaConsulta2.RecordSource = "SELECT SUM(Debe2) AS SumaDeDebe2, SUM(Haber2) AS SumaDeHaber2, SUM(Debe1) AS SumaDeDebe1, SUM(Haber1) AS SumaDeHaber1 From Reportes WHERE     (Descripcion LIKE 'Total%') AND (KeyGrupo = 'B' OR  KeyGrupo = 'C')"
   FrmReportes.DtaConsulta2.Refresh
   If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
      FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2,Reportes.Haber2,Reportes.KeyGrupo From Reportes Where (((Reportes.KeyGrupo) = 'PC'))"
      FrmReportes.DtaConsulta.Refresh
      If Not FrmReportes.DtaConsulta.Recordset.EOF Then
       If Not IsNull(FrmReportes.DtaConsulta2.Recordset("SumaDeHaber2")) Then
        TotalCuenta = FrmReportes.DtaConsulta2.Recordset("SumaDeHaber2")
         FrmReportes.DtaConsulta.Recordset("Haber1") = FrmReportes.DtaConsulta2.Recordset("SumaDeHaber1")
         FrmReportes.DtaConsulta.Recordset("Haber2") = TotalCuenta
        FrmReportes.DtaConsulta.Recordset.Update
       End If
      End If
   End If
   
 ElseIf QUIEN = "Resultado" Then
 FrmReportes.DtaConsulta2.RecordSource = "SELECT SUM(Debe2) AS SumaDeDebe2, SUM(Haber2) AS SumaDeHaber2, SUM(Debe1) AS SumaDeDebe1, SUM(Haber1) AS SumaDeHaber1 From Reportes WHERE     (Descripcion LIKE 'Total%') AND((Reportes.KeyGrupo) = 'G' Or (Reportes.KeyGrupo) = 'O')"
    'FrmReportes.DtaConsulta2.RecordSource = "SELECT Sum(Reportes.Debe1) AS SumaDeDebe1, Sum(Reportes.Haber1) AS SumaDeHaber1 From Reportes Where (((Reportes.Descripcion) Like 'Total%') And ((Reportes.KeyGrupo) = 'G' Or (Reportes.KeyGrupo) = 'O'))"
'   InputBox "", "", FrmReportes.DtaConsulta2.RecordSource
   FrmReportes.DtaConsulta2.Refresh
   If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
      FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2,Reportes.Haber2, Reportes.KeyGrupo From Reportes Where (((Reportes.KeyGrupo) = 'CG'))"
'      FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2,Reportes.Haber2,Reportes.KeyGrupo From Reportes Where (((Reportes.KeyGrupo) = 'CG'))"
      FrmReportes.DtaConsulta.Refresh
      If Not FrmReportes.DtaConsulta.Recordset.EOF Then
       If Not IsNull(FrmReportes.DtaConsulta2.Recordset("SumaDeHaber2")) Then
        TotalCuenta = FrmReportes.DtaConsulta2.Recordset("SumaDeHaber2")
         FrmReportes.DtaConsulta.Recordset("Haber1") = FrmReportes.DtaConsulta2.Recordset("SumaDeHaber1")
         FrmReportes.DtaConsulta.Recordset("Haber2") = TotalCuenta
        FrmReportes.DtaConsulta.Recordset.Update
       End If
      End If
   End If
 End If
 
 
    '////////////////////////////////////////////////////////////////////////
     '////////////AGREGO LOS SALDOS ACUMULADOS ////////////////////
    '////////////////////////////////////////////////////////////////////////
    
'   FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.Debe1+Reportes.Debe2 AS TotalDebe, Reportes.Haber1+Reportes.Haber2 AS TotalHaber From Reportes"
'   FrmReportes.DtaConsulta2.Refresh
'
'   FrmReportes.lblProgreso.Caption = "Agregando Saldos Acumulados a las Cuentas"
'    FrmReportes.osProgress1.Value = 0
'    FrmReportes.osProgress1.Max = FrmReportes.DtaConsulta2.Recordset.RecordCount
'
'
'    'para evitar este problema se puede hacer una actualización sencilla y rápida con un query
'    'Update Reportes set debe3=debe1+debe2,haber3=haber1+haber2
'
'  Do While Not FrmReportes.DtaConsulta2.Recordset.EOF
'      FrmReportes.osProgress1.Value = FrmReportes.osProgress1.Value + 1
'      FrmReportes.DtaConsulta2.Recordset("Debe3") = FrmReportes.DtaConsulta2.Recordset("TotalDebe")
'      FrmReportes.DtaConsulta2.Recordset("Haber3") = FrmReportes.DtaConsulta2.Recordset("TotalHaber")
'
'    FrmReportes.DtaConsulta2.Recordset.Update
'   FrmReportes.DtaConsulta2.Recordset.MoveNext
'   Loop
'
'   FrmReportes.DtaConsulta2.Refresh
Ejecutar.Execute "Update Reportes set debe3=debe1+debe2,haber3=haber1+haber2"

End Sub


Public Sub LlenarDataCombos(ado As Adodc, ByRef DC As DataCombo, Campo As String, BoundColumn As String)
    'Conectar los datacombos a los adodc correspondientes
    On Error GoTo err
    Set DC.RowSource = ado
    DC.DataField = Campo
    DC.ListField = Campo
    DC.BoundColumn = BoundColumn
err:
      Dim Msj
   If err.Number <> 0 Then
        Msj = "Error N° " & Str(err.Number) & " fue generado por " _
        & err.Source & Chr(13) & err.Description
        MsgBox Msj, , "Error", err.HelpFile, err.HelpContext
    End If
End Sub


Public Sub EstructuraCatalogo(QUIEN As String)
Dim CodigoGrupo As String, CantRegistros As Double
Dim Nivel As Integer, Longitud As Integer
Dim Mayor() As String, CodGrupo() As String, DescripcionBalance As String
Dim TotalMayor() As String, TotalDescripcion As String
Dim KeySuperior As String, NumeroHijos As Double, NumeroHijosTotales As Double
Dim DescripCuenta As String, DescripcionPadre As String, KeyUltimo As String
Dim TipoCuenta As String, TipoMoneda As String, i As Double, J As Double, k As Double, NumRegistros As Double

Dim Orden As Long
Orden = 1
If QUIEN = "Catalogo" Then
     FrmReportes.DtaHistorial.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.Child, Grupos.DescripcionGrupo From Grupos ORDER BY Grupos.KeyGrupo"
End If
Debug.Print FrmReportes.DtaHistorial.RecordSource
   FrmReportes.DtaHistorial.Refresh
   
   If Not FrmReportes.DtaHistorial.Recordset.EOF Then
    FrmReportes.DtaHistorial.Recordset.MoveLast
    CantRegistros = FrmReportes.DtaHistorial.Recordset.RecordCount
    FrmReportes.DtaHistorial.Recordset.MoveFirst
   End If
   
       FrmReportes.LblProgreso.Caption = "Creando Estructura"
       FrmReportes.osProgress1.Value = 0
       FrmReportes.osProgress1.Visible = True
       FrmReportes.osProgress1.Max = CantRegistros
       J = 0
   
     Do While Not FrmReportes.DtaHistorial.Recordset.EOF
      If Not IsNull(FrmReportes.DtaHistorial.Recordset("KeyGrupoSuperior")) Then
        KeySuperior = FrmReportes.DtaHistorial.Recordset("KeyGrupoSuperior")
      Else
        KeySuperior = ""
      End If
      CodigoGrupo = FrmReportes.DtaHistorial.Recordset("KeyGrupo")
      TotalDescripcion = "Total " + FrmReportes.DtaHistorial.Recordset("DescripcionGrupo")
      Descripcion = FrmReportes.DtaHistorial.Recordset("DescripcionGrupo")
      
      DoEvents
      
'//////////////////////IDENTIFICO EL NIVEL////////////////////////////////////////////////////
      Longitud = Len(FrmReportes.DtaHistorial.Recordset("KeyGrupo"))
      If Longitud > 1 Then
       If Longitud = 5 Then
        Nivel = 2
       Else
        Nivel = (Longitud - 5) / 2
        Nivel = Nivel + 2
       End If
      Else
       Nivel = 1
      End If
      
'////////////////////Lleno de Espacios la Descripcion del Grupo///////////////////////////////////
       For i = 2 To Nivel
        Descripcion = " " + Descripcion
        TotalDescripcion = " " + TotalDescripcion
       Next
'/////////////////////AGREGO EL GRUPO AL REPORTE////////////////////////////////////////////////

           FrmReportes.DtaReportes.Recordset.AddNew
              FrmReportes.DtaReportes.Recordset("Descripcion") = Descripcion
              FrmReportes.DtaReportes.Recordset("KeyGrupo") = CodigoGrupo
              FrmReportes.DtaReportes.Recordset("Nivel") = Nivel
              If Not KeySuperior = "" Then
                FrmReportes.DtaReportes.Recordset("KeyGrupoSuperior") = KeySuperior
              End If
              FrmReportes.DtaReportes.Recordset!Orden = Orden
            FrmReportes.DtaReportes.Recordset.Update
            Orden = Orden + 1
    

      
      
'///////////////////////Busco si Existen Cuentas para esteGrupo////////////////////////////////
       'FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas Where (((Transacciones.FechaTransaccion) <= " & NumFecha2 & ")) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo Having (((Cuentas.KeyGrupo) = '" & CodigoGrupo & "')) ORDER BY Cuentas.KeyGrupo"
       FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.TipoCuenta,Cuentas.TipoMoneda, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, Cuentas.DescripcionCuentas From Cuentas Where (((Cuentas.KeyGrupo) = '" & CodigoGrupo & "')) ORDER BY Cuentas.CodCuentas, Cuentas.KeyGrupo"
       FrmReportes.DtaConsulta.Refresh
       
       If Not FrmReportes.DtaConsulta.Recordset.EOF Then
        FrmReportes.DtaConsulta.Recordset.MoveLast
        NumRegistros = FrmReportes.DtaConsulta.Recordset.RecordCount
        FrmReportes.DtaConsulta.Recordset.MoveFirst
       End If
       
       FrmReportes.osProgress2.Value = 0
       FrmReportes.osProgress2.Visible = True
       FrmReportes.osProgress2.Max = NumRegistros
       k = 0
       
       Do While Not FrmReportes.DtaConsulta.Recordset.EOF
       TipoMoneda = FrmReportes.DtaConsulta.Recordset("TipoMoneda")
       TipoCuenta = FrmReportes.DtaConsulta.Recordset("TipoCuenta")
       
       
       
       If Not IsNull(FrmReportes.DtaConsulta.Recordset("DescripcionCuentas")) Then
        DescripCuenta = FrmReportes.DtaConsulta.Recordset("CodCuentas") + "." + FrmReportes.DtaConsulta.Recordset("DescripcionCuentas")
       Else
        DescripCuenta = "La Cuenta no tiene descripcion"
       End If
       CodigoCuenta = FrmReportes.DtaConsulta.Recordset("CodCuentas")
       
       FrmReportes.LblProgreso.Caption = "Agregando la Cuenta " & CodigoCuenta
       DoEvents
       
'/////////////Lleno de Espacios el codigo de la cuenta//////////////////////////////
           For i = 1 To Nivel
             DescripCuenta = " " + DescripCuenta
           Next
           FrmReportes.DtaReportes.Refresh
           FrmReportes.DtaReportes.Recordset.AddNew
              FrmReportes.DtaReportes.Recordset("Descripcion") = DescripCuenta
              FrmReportes.DtaReportes.Recordset("KeyGrupo") = CodigoCuenta
              FrmReportes.DtaReportes.Recordset("KeyGrupoCuenta") = CodigoGrupo
              FrmReportes.DtaReportes.Recordset("Nivel") = Nivel + 1
              FrmReportes.DtaReportes.Recordset!Orden = Orden
            FrmReportes.DtaReportes.Recordset.Update
            Orden = Orden + 1
            
            k = k + 1
            FrmReportes.osProgress2.Value = k
           FrmReportes.DtaConsulta.Recordset.MoveNext
         Loop
       
       J = J + 1
       FrmReportes.osProgress1.Value = J
       DoEvents
       FrmReportes.DtaHistorial.Recordset.MoveNext
       
     Loop

'FrmReportes.osProgress1.Visible = False
'FrmReportes.osProgress2.Visible = False

End Sub
Public Sub EliminaRegistroCeroDpto(QUIEN As String)
Dim KeyGrupo As String, Niveles As Integer
Dim rs As New ADODB.Recordset, cadena As String

If QUIEN = "Balanza" Then
'FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.Debe3-Reportes.Haber3 AS Diferencia From Reportes Where (((Reportes.Descripcion) Like '*Total*') And ((Reportes.Debe3 - Reportes.Haber3) = 0))"
FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.Debe3-Reportes.Haber3 AS Diferencia From Reportes Where (((Reportes.Descripcion) Like '%Total%') And ((Reportes.Debe1 + Reportes.Debe2+ Reportes.Debe3 + Reportes.Haber1+Reportes.Haber2+Reportes.Haber3) = 0))"
FrmReportes.DtaConsulta.Refresh
 If Not FrmReportes.DtaConsulta.Recordset.EOF Then
   FrmReportes.DtaConsulta.Recordset.MoveLast
   numero = FrmReportes.DtaConsulta.Recordset.RecordCount
   FrmReportes.DtaConsulta.Recordset.MoveFirst
   Do While Not FrmReportes.DtaConsulta.Recordset.EOF
      Descripcion = FrmReportes.DtaConsulta.Recordset("Descripcion")
      KeyGrupo = FrmReportes.DtaConsulta.Recordset("KeyGrupo")
      FrmReportes.DtaConsulta.Recordset.Delete
      FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta From Reportes Where (((Reportes.KeyGrupo) = '" & KeyGrupo & "'))"
      FrmReportes.DtaConsulta2.Refresh
      If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
         FrmReportes.DtaConsulta2.Recordset.Delete
      End If
  
     FrmReportes.DtaConsulta.Recordset.MoveNext
   Loop
   
   
   
 End If
ElseIf QUIEN = "Nivel" Then
  Niveles = Val(FrmReportes.CmbNivel.Text)
  FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta, Reportes.Nivel From Reportes Where (((Reportes.Nivel) > " & Niveles & "))"
  FrmReportes.DtaConsulta.Refresh
  If Not FrmReportes.DtaConsulta.Recordset.EOF Then
     FrmReportes.DtaConsulta.Recordset.MoveLast
     numero = FrmReportes.DtaConsulta.Recordset.RecordCount
     FrmReportes.DtaConsulta.Recordset.MoveFirst
'     Do While Not FrmReportes.DtaConsulta.Recordset.EOF
'        FrmReportes.DtaConsulta.Recordset.Delete
'        FrmReportes.DtaConsulta.Recordset.MoveNext
'     Loop

'///////////////////////////ELIMINO TODOS LOS REGISTROS MAYORES AL NIVEL SELECCIONADO ////////////////////
      rs.Open "DELETE FROM Reportes WHERE (((Reportes.Nivel) > " & Niveles & "))", Conexion
     FrmReportes.DtaConsulta.Refresh
  End If
  
' FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.Debe3-Reportes.Haber3 AS Diferencia From Reportes Where (((Reportes.Descripcion) Like '*Total*') And ((Reportes.Debe3 - Reportes.Haber3) = 0))"
FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.Debe3-Reportes.Haber3 AS Diferencia From Reportes Where (((Reportes.Descripcion) Like '%Total%') And ((Reportes.Debe3 - Reportes.Haber3) = 0)) ORDER BY Descripcion"
 FrmReportes.DtaConsulta.Refresh
 If Not FrmReportes.DtaConsulta.Recordset.EOF Then
   FrmReportes.DtaConsulta.Recordset.MoveLast
   numero = FrmReportes.DtaConsulta.Recordset.RecordCount
   FrmReportes.DtaConsulta.Recordset.MoveFirst
   Do While Not FrmReportes.DtaConsulta.Recordset.EOF
     Descripcion = Trim(FrmReportes.DtaConsulta.Recordset("Descripcion"))
     If Mid(Descripcion, 1, 5) = "Total" Then
       Descripcion = Mid(Descripcion, 6, Len(Descripcion))
     End If
     FrmReportes.Caption = Descripcion
     DoEvents
     KeyGrupo = FrmReportes.DtaConsulta.Recordset("KeyGrupo")
'     FrmReportes.DtaConsulta.Recordset.Delete
''     FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta From Reportes Where (((Reportes.KeyGrupo) = '" & KeyGrupo & "'))"
'     FrmReportes.DtaConsulta2.RecordSource = "SELECT Descripcion, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta From Reportes WHERE (KeyGrupo = '" & KeyGrupo & "') AND (Descripcion LIKE '%" & Descripcion & "%')"
'     FrmReportes.DtaConsulta2.Refresh
'     If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
'       FrmReportes.DtaConsulta2.Recordset.Delete
'     End If
     
     rs.Open "DELETE FROM Reportes WHERE (Descripcion LIKE '%" & Descripcion & "%') AND (KeyGrupo = '" & KeyGrupo & "') And ((Reportes.Debe3 - Reportes.Haber3) = 0)", Conexion

     
    FrmReportes.DtaConsulta.Recordset.MoveNext
   Loop
   
  FrmReportes.DtaConsulta.Refresh
 End If
Else


FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.Debe3-Reportes.Haber3 AS Diferencia From Reportes Where (((Reportes.Descripcion) Like '%Total%') And ((Reportes.Debe3 - Reportes.Haber3) = 0))"
FrmReportes.DtaConsulta.Refresh
 If Not FrmReportes.DtaConsulta.Recordset.EOF Then
  FrmReportes.DtaConsulta.Recordset.MoveLast
  numero = FrmReportes.DtaConsulta.Recordset.RecordCount
  FrmReportes.DtaConsulta.Refresh
  Do While Not FrmReportes.DtaConsulta.Recordset.EOF
    Descripcion = FrmReportes.DtaConsulta.Recordset("Descripcion")
    KeyGrupo = FrmReportes.DtaConsulta.Recordset("KeyGrupo")
    FrmReportes.DtaConsulta.Recordset.Delete
'    FrmReportes.DtaConsulta.Refresh
    FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta From Reportes Where (((Reportes.KeyGrupo) = '" & KeyGrupo & "'))"
    FrmReportes.DtaConsulta2.Refresh
    If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
      
     rs.Open "DELETE FROM Reportes WHERE (KeyGrupo = '" & KeyGrupo & "')", Conexion
    End If
    

    FrmReportes.DtaConsulta.Recordset.MoveNext
  Loop
 End If
End If
End Sub


Public Sub EliminaRegistroCero(QUIEN As String)
Dim KeyGrupo As String, Niveles As Integer
Dim rs As New ADODB.Recordset, cadena As String

If QUIEN = "Balanza" Then
'FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.Debe3-Reportes.Haber3 AS Diferencia From Reportes Where (((Reportes.Descripcion) Like '*Total*') And ((Reportes.Debe3 - Reportes.Haber3) = 0))"
FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.Debe3-Reportes.Haber3 AS Diferencia From Reportes Where (((Reportes.Descripcion) Like '%Total%') And ((Reportes.Debe1 + Reportes.Debe2+ Reportes.Debe3 + Reportes.Haber1+Reportes.Haber2+Reportes.Haber3) = 0))"
FrmReportes.DtaConsulta.Refresh
 If Not FrmReportes.DtaConsulta.Recordset.EOF Then
   FrmReportes.DtaConsulta.Recordset.MoveLast
   numero = FrmReportes.DtaConsulta.Recordset.RecordCount
   FrmReportes.DtaConsulta.Recordset.MoveFirst
   Do While Not FrmReportes.DtaConsulta.Recordset.EOF
      Descripcion = FrmReportes.DtaConsulta.Recordset("Descripcion")
      KeyGrupo = FrmReportes.DtaConsulta.Recordset("KeyGrupo")
      FrmReportes.DtaConsulta.Recordset.Delete
      FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta From Reportes Where (((Reportes.KeyGrupo) = '" & KeyGrupo & "'))"
      FrmReportes.DtaConsulta2.Refresh
      If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
         FrmReportes.DtaConsulta2.Recordset.Delete
      End If
  
     FrmReportes.DtaConsulta.Recordset.MoveNext
   Loop
   
   
   
 End If
ElseIf QUIEN = "Nivel" Then
  Niveles = Val(FrmReportes.CmbNivel.Text)
  FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta, Reportes.Nivel From Reportes Where (((Reportes.Nivel) > " & Niveles & "))"
  FrmReportes.DtaConsulta.Refresh
  If Not FrmReportes.DtaConsulta.Recordset.EOF Then
     FrmReportes.DtaConsulta.Recordset.MoveLast
     numero = FrmReportes.DtaConsulta.Recordset.RecordCount
     FrmReportes.DtaConsulta.Recordset.MoveFirst
     Do While Not FrmReportes.DtaConsulta.Recordset.EOF
        FrmReportes.DtaConsulta.Recordset.Delete
        FrmReportes.DtaConsulta.Recordset.MoveNext
     Loop
     FrmReportes.DtaConsulta.Refresh
  End If
  
' FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.Debe3-Reportes.Haber3 AS Diferencia From Reportes Where (((Reportes.Descripcion) Like '*Total*') And ((Reportes.Debe3 - Reportes.Haber3) = 0))"
FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.Debe3-Reportes.Haber3 AS Diferencia From Reportes Where (((Reportes.Descripcion) Like '%Total%') And ((Reportes.Debe3 - Reportes.Haber3) = 0))"
 FrmReportes.DtaConsulta.Refresh
 If Not FrmReportes.DtaConsulta.Recordset.EOF Then
   FrmReportes.DtaConsulta.Recordset.MoveLast
   numero = FrmReportes.DtaConsulta.Recordset.RecordCount
   FrmReportes.DtaConsulta.Recordset.MoveFirst
   Do While Not FrmReportes.DtaConsulta.Recordset.EOF
     Descripcion = Trim(FrmReportes.DtaConsulta.Recordset("Descripcion"))
     KeyGrupo = FrmReportes.DtaConsulta.Recordset("KeyGrupo")
     FrmReportes.DtaConsulta.Recordset.Delete
     FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta From Reportes Where (((Reportes.KeyGrupo) = '" & KeyGrupo & "'))"
     FrmReportes.DtaConsulta2.Refresh
     If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
       FrmReportes.DtaConsulta2.Recordset.Delete
     End If
  
    FrmReportes.DtaConsulta.Recordset.MoveNext
   Loop
  FrmReportes.DtaConsulta.Refresh
 End If
Else


FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.Debe3-Reportes.Haber3 AS Diferencia From Reportes Where (((Reportes.Descripcion) Like '%Total%') And ((Reportes.Debe3 - Reportes.Haber3) = 0))"
FrmReportes.DtaConsulta.Refresh
 If Not FrmReportes.DtaConsulta.Recordset.EOF Then
  FrmReportes.DtaConsulta.Recordset.MoveLast
  numero = FrmReportes.DtaConsulta.Recordset.RecordCount
  FrmReportes.DtaConsulta.Refresh
  Do While Not FrmReportes.DtaConsulta.Recordset.EOF
    Descripcion = FrmReportes.DtaConsulta.Recordset("Descripcion")
    KeyGrupo = FrmReportes.DtaConsulta.Recordset("KeyGrupo")
    FrmReportes.DtaConsulta.Recordset.Delete
'    FrmReportes.DtaConsulta.Refresh
    FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta From Reportes Where (((Reportes.KeyGrupo) = '" & KeyGrupo & "'))"
    FrmReportes.DtaConsulta2.Refresh
    If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
      
     rs.Open "DELETE FROM Reportes WHERE (KeyGrupo = '" & KeyGrupo & "')", Conexion
    End If
    

    FrmReportes.DtaConsulta.Recordset.MoveNext
  Loop
 End If
End If
End Sub


Public Sub CreaEstructura(QUIEN As String)
Dim CodigoGrupo As String, Orden As Integer
Dim Nivel As Integer, Longitud As Integer
Dim Mayor() As String, CodGrupo() As String, DescripcionBalance As String
Dim TotalMayor() As String, TotalDescripcion As String
Dim KeySuperior As String, NumeroHijos As Double, NumeroHijosTotales As Double
Dim DescripCuenta As String, DescripcionPadre As String, KeyUltimo As String
Dim UbicacionReporte As String, Fecha2 As String, CodigoCuentaDesde As String, CodigoCuentaHasta As String
Dim CodDepartamento As String, NPeriodo As Double, NumeroPeriodo() As Double, i As Double, ContPeriodo As Double
'///////////////////////////////////////////////////////////////////////
'//////////////////////Esta es la Estructura de la Balanza/////////////
'//////////////////////////////////////////////////////////////////////
 If QUIEN = "Balanza" Then
 
             If FrmReportes.TxtDesde.Text = "" Then
               FrmReportes.DtaConsulta.RecordSource = "SELECT * From Grupos ORDER BY KeyGrupo"
               FrmReportes.DtaConsulta.Refresh
               If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                 FrmReportes.DtaConsulta.Recordset.MoveFirst
                 CodigoCuentaDesde = FrmReportes.DtaConsulta.Recordset("KeyGrupo")
               End If
            Else
                CodigoCuentaDesde = FrmReportes.TxtKeyGrupoDesde.Text
            End If
               
            If FrmReportes.TxtHasta.Text = "" Then
               FrmReportes.DtaConsulta.RecordSource = "SELECT * From Grupos ORDER BY KeyGrupo"
               FrmReportes.DtaConsulta.Refresh
               If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                 FrmReportes.DtaConsulta.Recordset.MoveLast
                 CodigoCuentaHasta = FrmReportes.DtaConsulta.Recordset("KeyGrupo")
               End If
            Else
               CodigoCuentaHasta = FrmReportes.TxtKeyGrupoHasta.Text
            End If
 

     FrmReportes.DtaHistorial.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.Child, Grupos.DescripcionGrupo From Grupos WHERE (KeyGrupo BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') ORDER BY KeyGrupo"
 ElseIf QUIEN = "Balance" Then
     FrmReportes.DtaHistorial.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.Child, Grupos.DescripcionGrupo From Grupos Where (((Grupos.KeyGrupo) Like 'A%' Or (Grupos.KeyGrupo) Like 'B%' Or (Grupos.KeyGrupo) Like 'C%')) ORDER BY Grupos.KeyGrupo"
 ElseIf QUIEN = "Resultado" Then
     FrmReportes.DtaHistorial.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.Child, Grupos.DescripcionGrupo From Grupos Where (((Grupos.KeyGrupo) Like 'D%' Or (Grupos.KeyGrupo) Like 'G%' Or (Grupos.KeyGrupo) Like 'O%')) ORDER BY Grupos.KeyGrupo"
 ElseIf QUIEN = "BalanzaLibroMayor" Then
   
            If FrmReportes.TxtDesde.Text = "" Then
               FrmReportes.DtaConsulta.RecordSource = "SELECT * From Grupos ORDER BY KeyGrupo"
               FrmReportes.DtaConsulta.Refresh
               If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                 FrmReportes.DtaConsulta.Recordset.MoveFirst
                 CodigoCuentaDesde = FrmReportes.DtaConsulta.Recordset("KeyGrupo")
               End If
            Else
                CodigoCuentaDesde = FrmReportes.TxtKeyGrupoDesde.Text
            End If
               
            If FrmReportes.TxtHasta.Text = "" Then
               FrmReportes.DtaConsulta.RecordSource = "SELECT * From Grupos ORDER BY KeyGrupo"
               FrmReportes.DtaConsulta.Refresh
               If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                 FrmReportes.DtaConsulta.Recordset.MoveLast
                 CodigoCuentaHasta = FrmReportes.DtaConsulta.Recordset("KeyGrupo")
               End If
            Else
               CodigoCuentaHasta = FrmReportes.TxtKeyGrupoHasta.Text
            End If
            
            FrmReportes.DtaConsulta.RecordSource = "SELECT NPeriodo, NumeroTabla, FechaPeriodo, EstadoPeriodo, NTransacciones, Periodo From Periodos WHERE (FechaPeriodo BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) "
            FrmReportes.DtaConsulta.Refresh
            If Not FrmReportes.DtaConsulta.Recordset.EOF Then
               FrmReportes.DtaConsulta.Recordset.MoveLast
               ContPeriodo = FrmReportes.DtaConsulta.Recordset.RecordCount
               ReDim NumeroPeriodo(ContPeriodo) As Double
               FrmReportes.DtaConsulta.Recordset.MoveFirst
            End If
            
            i = 1
            Do While Not FrmReportes.DtaConsulta.Recordset.EOF
               NumeroPeriodo(i) = FrmReportes.DtaConsulta.Recordset("NPeriodo")
              i = i + 1
              FrmReportes.DtaConsulta.Recordset.MoveNext
            Loop
            
            FrmReportes.DtaHistorial.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.Child, Grupos.DescripcionGrupo From Grupos WHERE (KeyGrupo BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') ORDER BY KeyGrupo"
 
 End If
    Orden = 0
     FrmReportes.DtaHistorial.Refresh
'     InputBox "", "", FrmReportes.DtaHistorial.RecordSource
        FrmReportes.LblProgreso.Caption = "Creando Estructura"
        FrmReportes.osProgress1.Value = 0
        FrmReportes.osProgress1.Visible = True
       FrmReportes.osProgress1.Max = FrmReportes.DtaHistorial.Recordset.RecordCount
       
       
     Do While Not FrmReportes.DtaHistorial.Recordset.EOF
        FrmReportes.osProgress1.Value = FrmReportes.osProgress1.Value + 1
      If Not IsNull(FrmReportes.DtaHistorial.Recordset("KeyGrupoSuperior")) Then
        KeySuperior = FrmReportes.DtaHistorial.Recordset("KeyGrupoSuperior")
      Else
        KeySuperior = ""
      End If
      CodigoGrupo = FrmReportes.DtaHistorial.Recordset("KeyGrupo")
      TotalDescripcion = "Total " + FrmReportes.DtaHistorial.Recordset("DescripcionGrupo")
      Descripcion = FrmReportes.DtaHistorial.Recordset("DescripcionGrupo")

      
'//////////////////////IDENTIFICO EL NIVEL////////////////////////////////////////////////////
      Longitud = Len(FrmReportes.DtaHistorial.Recordset("KeyGrupo"))
      If Longitud > 1 Then
       If Longitud = 5 Then
        Nivel = 2
       Else
        Nivel = (Longitud - 5) / 2
        Nivel = Nivel + 2
       End If
      Else
       Nivel = 1
      End If
      
'////////////////////Lleno de Espacios la Descripcion del Grupo///////////////////////////////////
       For i = 2 To Nivel
        Descripcion = " " + Descripcion
        TotalDescripcion = " " + TotalDescripcion
       Next
'/////////////////////AGREGO EL GRUPO AL REPORTE////////////////////////////////////////////////
           Orden = Orden + 1
           FrmReportes.DtaReportes.Refresh
'           FrmReportes.DtaReportes.Recordset.MoveLast
           FrmReportes.DtaReportes.Recordset.AddNew
              FrmReportes.DtaReportes.Recordset("Descripcion") = Descripcion
              FrmReportes.DtaReportes.Recordset("KeyGrupo") = CodigoGrupo
              FrmReportes.DtaReportes.Recordset("Nivel") = Nivel
              FrmReportes.DtaReportes.Recordset("Orden") = Orden
              If Not KeySuperior = "" Then
                FrmReportes.DtaReportes.Recordset("KeyGrupoSuperior") = KeySuperior
              End If
            FrmReportes.DtaReportes.Recordset.Update
    

      
       NumFecha2 = FechaFin
       Fecha2 = Format(FechaFin, "yyyy-mm-dd")
       Fecha1 = Format(FechaIni, "yyyy-mm-dd")
'///////////////////////Busco si Existen Cuentas para esteGrupo////////////////////////////////

          FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo,Cuentas.UbicacionReporte FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas Where  (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo,Cuentas.UbicacionReporte Having (((Cuentas.KeyGrupo) = '" & CodigoGrupo & "')) ORDER BY Cuentas.KeyGrupo"
          FrmReportes.DtaConsulta.Refresh
         


       'Transacciones.VoucherNo
      
       NPeriodo = 0   ' ESTO SOLO FUNCIONA PARA LOS LIBROS MAYOR
       
       Do While Not FrmReportes.DtaConsulta.Recordset.EOF
       If Not IsNull(FrmReportes.DtaConsulta.Recordset("DescripcionCuentas")) Then
        DescripCuenta = FrmReportes.DtaConsulta.Recordset("CodCuentas") + "." + FrmReportes.DtaConsulta.Recordset("DescripcionCuentas")
       Else
        DescripCuenta = FrmReportes.DtaConsulta.Recordset("CodCuentas") + "." + "NO TIENE DESCRIPCION????"
       End If
       CodigoCuenta = FrmReportes.DtaConsulta.Recordset("CodCuentas")
'       If Not IsNull(FrmReportes.DtaConsulta.Recordset("VoucherNo")) Then
'        CodDepartamento = FrmReportes.DtaConsulta.Recordset("VoucherNo")
'       End If

       If Not IsNull(FrmReportes.DtaConsulta.Recordset("UbicacionReporte")) Then
        UbicacionReporte = FrmReportes.DtaConsulta.Recordset("UbicacionReporte")
       Else
          Select Case FrmReportes.DtaConsulta.Recordset("TipoCuenta")
            Case "Caja": UbicacionReporte = "Cajas"
            Case "Bancos": UbicacionReporte = "Bancos"
            Case "Cuentas x Cobrar": UbicacionReporte = "Cuentas x Cobrar"
            Case "Inventario": UbicacionReporte = "Inventario"
            Case "Activo Fijo": UbicacionReporte = "Terreno y Edificios"
            Case "Papeleria - Utiles": UbicacionReporte = "Papeleria y Utiles de Oficina"
            Case "Otros Activos": UbicacionReporte = "Otros Activos"
            Case "Cuentas x Pagar": UbicacionReporte = "Proveedores"
            Case "Pasivo": UbicacionReporte = "Pasivos Acumulados"
            Case "Otros Pasivos": UbicacionReporte = "Otros Pasivos"
            Case "Capital": UbicacionReporte = "Acciones Comunes"
            Case "Ingresos - Ventas": UbicacionReporte = "Ingresos - Ventas"
            Case "Costos": UbicacionReporte = "Costos"
            Case "Gastos": UbicacionReporte = "Gastos"
            
        End Select
       End If
       
       If UbicacionReporte = " " Then
         Select Case FrmReportes.DtaConsulta.Recordset("TipoCuenta")
            Case "Caja": UbicacionReporte = "Cajas"
            Case "Bancos": UbicacionReporte = "Bancos"
            Case "Cuentas x Cobrar": UbicacionReporte = "Cuentas x Cobrar"
            Case "Inventario": UbicacionReporte = "Inventario"
            Case "Activo Fijo": UbicacionReporte = "Terreno y Edificios"
            Case "Papeleria - Utiles": UbicacionReporte = "Papeleria y Utiles de Oficina"
            Case "Otros Activos": UbicacionReporte = "Otros Activos"
            Case "Cuentas x Pagar": UbicacionReporte = "Proveedores"
            Case "Pasivo": UbicacionReporte = "Pasivos Acumulados"
            Case "Otros Pasivos": UbicacionReporte = "Otros Pasivos"
            Case "Capital": UbicacionReporte = "Acciones Comunes"
            Case "Ingresos - Ventas": UbicacionReporte = "Ingresos - Ventas"
            Case "Costos": UbicacionReporte = "Costos"
            Case "Gastos": UbicacionReporte = "Gastos"
            
        End Select
       ElseIf UbicacionReporte = "" Then
         Select Case FrmReportes.DtaConsulta.Recordset("TipoCuenta")
            Case "Caja": UbicacionReporte = "Cajas"
            Case "Bancos": UbicacionReporte = "Bancos"
            Case "Cuentas x Cobrar": UbicacionReporte = "Cuentas x Cobrar"
            Case "Inventario": UbicacionReporte = "Inventario"
            Case "Activo Fijo": UbicacionReporte = "Terreno y Edificios"
            Case "Papeleria - Utiles": UbicacionReporte = "Papeleria y Utiles de Oficina"
            Case "Otros Activos": UbicacionReporte = "Otros Activos"
            Case "Cuentas x Pagar": UbicacionReporte = "Proveedores"
            Case "Pasivo": UbicacionReporte = "Pasivos Acumulados"
            Case "Otros Pasivos": UbicacionReporte = "Otros Pasivos"
            Case "Capital": UbicacionReporte = "Acciones Comunes"
            Case "Ingresos - Ventas": UbicacionReporte = "Ingresos - Ventas"
            Case "Costos": UbicacionReporte = "Costos"
            Case "Gastos": UbicacionReporte = "Gastos"
            
        End Select
      End If

       
'/////////////Lleno de Espacios el codigo de la cuenta//////////////////////////////
           For i = 1 To Nivel
             DescripCuenta = " " + DescripCuenta
           Next
           
           
            Orden = Orden + 1

              If QUIEN = "BalanzaLibroMayor" Then
                For i = 1 To ContPeriodo
                    FrmReportes.DtaReportes.Recordset.AddNew
                    FrmReportes.DtaReportes.Recordset("Descripcion") = DescripCuenta
                    FrmReportes.DtaReportes.Recordset("KeyGrupo") = CodigoCuenta
                    FrmReportes.DtaReportes.Recordset("KeyGrupoCuenta") = CodigoGrupo
                    FrmReportes.DtaReportes.Recordset("Nivel") = Nivel + 1
                    FrmReportes.DtaReportes.Recordset("Ubicacion") = UbicacionReporte
                    FrmReportes.DtaReportes.Recordset("Orden") = Orden
                    FrmReportes.DtaReportes.Recordset("CodCuentas") = CodigoCuenta
                    FrmReportes.DtaReportes.Recordset("Nperiodo") = NumeroPeriodo(i)
                    FrmReportes.DtaReportes.Recordset.Update
               Next
              

              Else
                FrmReportes.DtaReportes.Recordset.AddNew
                    FrmReportes.DtaReportes.Recordset("Descripcion") = DescripCuenta
                    FrmReportes.DtaReportes.Recordset("KeyGrupo") = CodigoCuenta
                    FrmReportes.DtaReportes.Recordset("KeyGrupoCuenta") = CodigoGrupo
                    FrmReportes.DtaReportes.Recordset("Nivel") = Nivel + 1
                    FrmReportes.DtaReportes.Recordset("Ubicacion") = UbicacionReporte
                    FrmReportes.DtaReportes.Recordset("Orden") = Orden
                    FrmReportes.DtaReportes.Recordset("CodCuentas") = CodigoCuenta
                FrmReportes.DtaReportes.Recordset.Update
              End If
            
            
            
            
           FrmReportes.DtaConsulta.Recordset.MoveNext
       Loop
  

'///////////Busco si este Grupo no tiene SubGrupos para totalizar/////////////////////////////////////////////
       FrmReportes.DtaConsulta.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.DescripcionGrupo From Grupos Where (((Grupos.KeyGrupoSuperior) = '" & CodigoGrupo & "'))"
       FrmReportes.DtaConsulta.Refresh
       If FrmReportes.DtaConsulta.Recordset.EOF Then
'/////////////////////////////Verifico si es de Balance, para Agregar Capital/////
          If QUIEN = "Balance" Then
            If CodigoGrupo = "C" Then
             DescripcionBalance = " Resultado Periodo"

               DescripcionBalance = " " + DescripcionBalance
               
              Orden = Orden + 1
              FrmReportes.DtaReportes.Recordset.AddNew
               FrmReportes.DtaReportes.Recordset("Descripcion") = DescripcionBalance
               FrmReportes.DtaReportes.Recordset("KeyGrupo") = CodigoGrupo
               FrmReportes.DtaReportes.Recordset("Nivel") = Nivel
              FrmReportes.DtaReportes.Recordset("Orden") = Orden
              FrmReportes.DtaReportes.Recordset("Ubicacion") = "Resultado Periodo"
              FrmReportes.DtaReportes.Recordset.Update
              
 
 
            End If
          End If
            Orden = Orden + 1
            FrmReportes.DtaReportes.Refresh
            FrmReportes.DtaReportes.Recordset.AddNew
              FrmReportes.DtaReportes.Recordset("Descripcion") = TotalDescripcion
              FrmReportes.DtaReportes.Recordset("KeyGrupo") = CodigoGrupo
              FrmReportes.DtaReportes.Recordset("Nivel") = Nivel
              FrmReportes.DtaReportes.Recordset("Orden") = Orden
              If Not KeySuperior = "" Then
               FrmReportes.DtaReportes.Recordset("KeyGrupoSuperior") = KeySuperior
              End If
            FrmReportes.DtaReportes.Recordset.Update
       
       End If
       
'///////////Busco si el padre de este grupo ya totalizaron todos sus hijos/////////////////////////////////////////////
       FrmReportes.DtaReportes.Refresh
       FrmReportes.DtaConsulta.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.DescripcionGrupo From Grupos Where (((Grupos.KeyGrupoSuperior) = '" & KeySuperior & "'))"
       FrmReportes.DtaConsulta.Refresh
       If Not FrmReportes.DtaConsulta.Recordset.EOF Then
     
         FrmReportes.DtaConsulta.Recordset.MoveLast
         NumeroHijos = FrmReportes.DtaConsulta.Recordset.RecordCount
         FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.KeyGrupo From Reportes Where (((Reportes.KeyGrupoSuperior) = '" & KeySuperior & "'))"
         FrmReportes.DtaConsulta2.Refresh
         If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
          FrmReportes.DtaConsulta2.Recordset.MoveLast
          NumeroHijosTotales = Val(FrmReportes.DtaConsulta2.Recordset.RecordCount) / 2
         End If
         
'/////////////////////Verifico todos los hijos ya tienen totales////////////////////////////////////
         If NumeroHijos = NumeroHijosTotales Then
            FrmReportes.DtaGrupos.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.DescripcionGrupo From Grupos Where (((Grupos.KeyGrupo) = '" & KeySuperior & "'))"
            FrmReportes.DtaGrupos.Refresh
            If Not FrmReportes.DtaGrupos.Recordset.EOF Then
              DescripcionPadre = "Total " + FrmReportes.DtaGrupos.Recordset("DescripcionGrupo")
'//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
              Longitud = Len(FrmReportes.DtaGrupos.Recordset("KeyGrupo"))
              If Longitud > 1 Then
                 If Longitud = 5 Then
                     Nivel = 2
                 Else
                     Nivel = (Longitud - 5) / 2
                     Nivel = Nivel + 2
                 End If
              Else
                    Nivel = 1
              End If
'////////////////////Lleno de Espacios la Descripcion del padre///////////////////////////////////
             For i = 2 To Nivel
               DescripcionPadre = " " + DescripcionPadre
             Next

'/////////////////////////////Verifico si es de Balance, para Agregar Capital/////
          If QUIEN = "Balance" Then
            If KeySuperior = "C" Then
             DescripcionBalance = "Resultado Periodo"

               DescripcionBalance = " " + DescripcionBalance
              FrmReportes.DtaReportes.Refresh
              Orden = Orden + 1
              FrmReportes.DtaReportes.Recordset.AddNew
               FrmReportes.DtaReportes.Recordset("Descripcion") = DescripcionBalance
               FrmReportes.DtaReportes.Recordset("KeyGrupo") = KeySuperior
              FrmReportes.DtaReportes.Recordset("Nivel") = Nivel
              FrmReportes.DtaReportes.Recordset("Orden") = Orden
              FrmReportes.DtaReportes.Recordset("Ubicacion") = "Resultado Periodo"
              FrmReportes.DtaReportes.Recordset.Update
              

 
            End If
          End If
'/////////////////////pongo el total del padre////////////////////////////
              Orden = Orden + 1
              
               If QUIEN = "BalanzaLibroMayor" Then
                For i = 1 To ContPeriodo
                  FrmReportes.DtaReportes.Recordset.AddNew
                    FrmReportes.DtaReportes.Recordset("Descripcion") = DescripcionPadre
                    FrmReportes.DtaReportes.Recordset("KeyGrupo") = KeySuperior
                    FrmReportes.DtaReportes.Recordset("Nivel") = Nivel
                    FrmReportes.DtaReportes.Recordset("Orden") = Orden
                     If Not IsNull(FrmReportes.DtaGrupos.Recordset("KeyGrupoSuperior")) Then
                      FrmReportes.DtaReportes.Recordset("KeyGrupoSuperior") = FrmReportes.DtaGrupos.Recordset("KeyGrupoSuperior")
                     End If
                    FrmReportes.DtaReportes.Recordset("Nperiodo") = NumeroPeriodo(i)
                  FrmReportes.DtaReportes.Recordset.Update
                
                Next
               Else
                FrmReportes.DtaReportes.Recordset.AddNew
                     FrmReportes.DtaReportes.Recordset("Descripcion") = DescripcionPadre
                     FrmReportes.DtaReportes.Recordset("KeyGrupo") = KeySuperior
                    FrmReportes.DtaReportes.Recordset("Nivel") = Nivel
                    FrmReportes.DtaReportes.Recordset("Orden") = Orden
                     If Not IsNull(FrmReportes.DtaGrupos.Recordset("KeyGrupoSuperior")) Then
                      FrmReportes.DtaReportes.Recordset("KeyGrupoSuperior") = FrmReportes.DtaGrupos.Recordset("KeyGrupoSuperior")
                     End If
                FrmReportes.DtaReportes.Recordset.Update
               End If

              FrmReportes.DtaReportes.Refresh
 
 
 '/////////////////////////Busco Si ya totalizaron los otros padres//////////////////////
       If Not IsNull(FrmReportes.DtaGrupos.Recordset("KeyGrupoSuperior")) Then
          Respuesta = FrmReportes.DtaGrupos.Recordset("KeyGrupoSuperior")
       End If
       
       KeyUltimo = KeySuperior
       Nivel = Nivel - 1
       For i = Nivel To 1 Step -1
'///////Verifico si el Grupo Anterior es el ultimo Hijo del Padre acutal

        FrmReportes.DtaConsulta.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.DescripcionGrupo From Grupos Where (((Grupos.KeyGrupoSuperior) = '" & Respuesta & "'))"
        FrmReportes.DtaConsulta.Refresh
        If Not FrmReportes.DtaConsulta.Recordset.EOF Then
           FrmReportes.DtaConsulta.Recordset.MoveLast
           KeyHijo = FrmReportes.DtaConsulta.Recordset("KeyGrupo")
           If KeyHijo = KeyUltimo Then
'////////Si es el ultimo Hijo Busco los datos del padre para Cerrarlo////////////////////////////
              KeySuperior = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
              FrmReportes.DtaConsulta2.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.DescripcionGrupo From Grupos Where (((Grupos.KeyGrupo) = '" & KeySuperior & "'))"
              FrmReportes.DtaConsulta2.Refresh
              DescripcionPadre = "Total " + FrmReportes.DtaConsulta2.Recordset("DescripcionGrupo")

'////////////////////Lleno de Espacios la Descripcion del padre///////////////////////////////////
             For J = 2 To i
               DescripcionPadre = " " + DescripcionPadre
             Next
           
'/////////////////////////////Verifico si es de Balance, para Agregar Capital/////
          If QUIEN = "Balance" Then
            If KeySuperior = "C" Then
             DescripcionBalance = "Resultado Periodo"

               DescripcionBalance = " " + DescripcionBalance
              Orden = Orden + 1
              FrmReportes.DtaReportes.Recordset.AddNew
               FrmReportes.DtaReportes.Recordset("Descripcion") = DescripcionBalance
               FrmReportes.DtaReportes.Recordset("KeyGrupo") = "RP"
              FrmReportes.DtaReportes.Recordset("Nivel") = Nivel
              FrmReportes.DtaReportes.Recordset("Orden") = Orden
              FrmReportes.DtaReportes.Recordset("Ubicacion") = "Resultado Periodo"
              FrmReportes.DtaReportes.Recordset.Update
         
            End If
          End If
'/////////////////////pongo el total del padre////////////////////////////
               Orden = Orden + 1
               FrmReportes.DtaReportes.Recordset.AddNew
               FrmReportes.DtaReportes.Recordset("Descripcion") = DescripcionPadre
               FrmReportes.DtaReportes.Recordset("KeyGrupo") = KeySuperior
              FrmReportes.DtaReportes.Recordset("Nivel") = i
              FrmReportes.DtaReportes.Recordset("Orden") = Orden
               If Not IsNull(FrmReportes.DtaConsulta2.Recordset("KeyGrupoSuperior")) Then
                FrmReportes.DtaReportes.Recordset("KeyGrupoSuperior") = FrmReportes.DtaConsulta2.Recordset("KeyGrupoSuperior")
               End If
              FrmReportes.DtaReportes.Recordset.Update
              KeyUltimo = KeySuperior
              FrmReportes.DtaReportes.Refresh
           
           Else
            Exit For
           End If
        
             

           If Not IsNull(FrmReportes.DtaConsulta2.Recordset("KeyGrupoSuperior")) Then
              Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupoSuperior")
           End If
        End If
       Next
            
            End If
         End If
       End If
       
       
       
       
       
       FrmReportes.DtaHistorial.Recordset.MoveNext
     Loop




End Sub
Public Sub CreaEstructuraDpto(QUIEN As String)
Dim CodigoGrupo As String, Orden As Integer
Dim Nivel As Integer, Longitud As Integer
Dim Mayor() As String, CodGrupo() As String, DescripcionBalance As String
Dim TotalMayor() As String, TotalDescripcion As String
Dim KeySuperior As String, NumeroHijos As Double, NumeroHijosTotales As Double
Dim DescripCuenta As String, DescripcionPadre As String, KeyUltimo As String
Dim UbicacionReporte As String, Fecha2 As String, CodigoCuentaDesde As String, CodigoCuentaHasta As String
Dim CodDepartamento As String, DescripcionDepartamento As String, TotalDescripcionDpto As String
Dim CodDepartamentoAnt As String, CodigoGrupoAnt As String, UltimoCodigoGrupo As String
'///////////////////////////////////////////////////////////////////////
'//////////////////////Esta es la Estructura de la Balanza/////////////
'//////////////////////////////////////////////////////////////////////
 If QUIEN = "Balanza" Then
 
             If FrmReportes.TxtDesde.Text = "" Then
               FrmReportes.DtaConsulta.RecordSource = "SELECT * From Grupos ORDER BY KeyGrupo"
               FrmReportes.DtaConsulta.Refresh
               If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                 FrmReportes.DtaConsulta.Recordset.MoveFirst
                 CodigoCuentaDesde = FrmReportes.DtaConsulta.Recordset("KeyGrupo")
               End If
            Else
                CodigoCuentaDesde = FrmReportes.TxtKeyGrupoDesde.Text
            End If
               
            If FrmReportes.TxtHasta.Text = "" Then
               FrmReportes.DtaConsulta.RecordSource = "SELECT * From Grupos ORDER BY KeyGrupo"
               FrmReportes.DtaConsulta.Refresh
               If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                 FrmReportes.DtaConsulta.Recordset.MoveLast
                 CodigoCuentaHasta = FrmReportes.DtaConsulta.Recordset("KeyGrupo")
               End If
            Else
               CodigoCuentaHasta = FrmReportes.TxtKeyGrupoHasta.Text
            End If
 

     FrmReportes.DtaHistorial.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.Child, Grupos.DescripcionGrupo From Grupos WHERE (KeyGrupo BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') ORDER BY KeyGrupo"
 ElseIf QUIEN = "Balance" Then
     FrmReportes.DtaHistorial.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.Child, Grupos.DescripcionGrupo From Grupos Where (((Grupos.KeyGrupo) Like 'A%' Or (Grupos.KeyGrupo) Like 'B%' Or (Grupos.KeyGrupo) Like 'C%')) ORDER BY Grupos.KeyGrupo"
 ElseIf QUIEN = "Resultado" Then
     FrmReportes.DtaHistorial.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.Child, Grupos.DescripcionGrupo From Grupos Where (((Grupos.KeyGrupo) Like 'D%' Or (Grupos.KeyGrupo) Like 'G%' Or (Grupos.KeyGrupo) Like 'O%')) ORDER BY Grupos.KeyGrupo"
  
 End If
 
 
    Orden = 0
     FrmReportes.DtaHistorial.Refresh
'     InputBox "", "", FrmReportes.DtaHistorial.RecordSource
        FrmReportes.LblProgreso.Caption = "Creando Estructura"
        FrmReportes.osProgress1.Value = 0
        FrmReportes.osProgress1.Visible = True
       FrmReportes.osProgress1.Max = FrmReportes.DtaHistorial.Recordset.RecordCount
       
       
     Do While Not FrmReportes.DtaHistorial.Recordset.EOF
        FrmReportes.osProgress1.Value = FrmReportes.osProgress1.Value + 1
      If Not IsNull(FrmReportes.DtaHistorial.Recordset("KeyGrupoSuperior")) Then
        KeySuperior = FrmReportes.DtaHistorial.Recordset("KeyGrupoSuperior")
      Else
        KeySuperior = ""
      End If
      CodigoGrupo = FrmReportes.DtaHistorial.Recordset("KeyGrupo")
      TotalDescripcion = "Total " + FrmReportes.DtaHistorial.Recordset("DescripcionGrupo")
      Descripcion = FrmReportes.DtaHistorial.Recordset("DescripcionGrupo")
      CodDepartamentoAnt = ""
      CodigoGrupoAnt = ""

      
'//////////////////////IDENTIFICO EL NIVEL////////////////////////////////////////////////////
      Longitud = Len(FrmReportes.DtaHistorial.Recordset("KeyGrupo"))
      If Longitud > 1 Then
       If Longitud = 5 Then
        Nivel = 2
       Else
        Nivel = (Longitud - 5) / 2
        Nivel = Nivel + 2
       End If
      Else
       Nivel = 1
      End If
      
'////////////////////Lleno de Espacios la Descripcion del Grupo///////////////////////////////////
       For i = 2 To Nivel
        Descripcion = " " + Descripcion
        TotalDescripcion = " " + TotalDescripcion
       Next
'/////////////////////AGREGO EL GRUPO AL REPORTE////////////////////////////////////////////////
           Orden = Orden + 1
           FrmReportes.DtaReportes.Refresh
'           FrmReportes.DtaReportes.Recordset.MoveLast
           FrmReportes.DtaReportes.Recordset.AddNew
              FrmReportes.DtaReportes.Recordset("Descripcion") = Descripcion
              FrmReportes.DtaReportes.Recordset("KeyGrupo") = CodigoGrupo
              FrmReportes.DtaReportes.Recordset("Nivel") = Nivel
              FrmReportes.DtaReportes.Recordset("Orden") = Orden
              If Not KeySuperior = "" Then
                FrmReportes.DtaReportes.Recordset("KeyGrupoSuperior") = KeySuperior
              End If
            FrmReportes.DtaReportes.Recordset.Update
    

      
       NumFecha2 = FechaFin
       Fecha2 = Format(FechaFin, "yyyy-mm-dd")
'///////////////////////Busco si Existen Cuentas para esteGrupo////////////////////////////////
       FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo,Cuentas.UbicacionReporte,Transacciones.VoucherNo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas Where  (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo,Cuentas.UbicacionReporte,Transacciones.VoucherNo Having (((Cuentas.KeyGrupo) = '" & CodigoGrupo & "')) ORDER BY Cuentas.KeyGrupo"
'       FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.Credito * Transacciones.TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, Cuentas.UbicacionReporte, Transacciones.VoucherNo, ISNULL(GrupoCuentas.CodGrupo,'SDpto') AS CodDepartamento, ISNULL(GrupoCuentas.DescripcionGrupo,'Sin Departamento') AS DescripcionDepartamento FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas LEFT OUTER JOIN GrupoCuentas ON Cuentas.CodGrupo = GrupoCuentas.CodGrupo " & _
'                                              "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, Cuentas.UbicacionReporte , Transacciones.VoucherNo, GrupoCuentas.CodGrupo, GrupoCuentas.DescripcionGrupo HAVING (Cuentas.KeyGrupo = '" & CodigoGrupo & "') ORDER BY Cuentas.KeyGrupo, CodDepartamento"
       FrmReportes.DtaConsulta.Refresh
       
       'Transacciones.VoucherNo
      
       
     Do While Not FrmReportes.DtaConsulta.Recordset.EOF
       If Not IsNull(FrmReportes.DtaConsulta.Recordset("DescripcionCuentas")) Then
        DescripCuenta = FrmReportes.DtaConsulta.Recordset("CodCuentas") + "." + FrmReportes.DtaConsulta.Recordset("DescripcionCuentas")
       Else
        DescripCuenta = FrmReportes.DtaConsulta.Recordset("CodCuentas") + "." + "NO TIENE DESCRIPCION????"
       End If
       CodigoCuenta = FrmReportes.DtaConsulta.Recordset("CodCuentas")
       If Not IsNull(FrmReportes.DtaConsulta.Recordset("VoucherNo")) Then
        CodDepartamento = FrmReportes.DtaConsulta.Recordset("VoucherNo")
       End If
       
       If CodDepartamento = "" Or CodDepartamento = "-" Then
          CodDepartamento = "-"
        End If
       
       
       '-----------------------------------------------------------------------------------------------------------------
       '----------------------------------------AGREGO LOS DEPARTAMENTOS ------------------------------------------------
       '-----------------------------------------------------------------------------------------------------------------
        DescripcionDepartamento = BuscaDpto(CodDepartamento)
        TotalDescripcionDpto = "Total " & DescripcionDepartamento
       
       '/////////////Lleno de Espacios el codigo de la cuenta//////////////////////////////
           For i = 1 To Nivel
             DescripcionDepartamento = " " + DescripcionDepartamento
             TotalDescripcionDpto = " " + TotalDescripcionDpto
           Next
           
         
           
           
        '/////////////////////AGREGO EL GRUPO AL REPORTE////////////////////////////////////////////////
           MDIPrimero.AdoConsulta.RecordSource = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, Nivel, CodCuentas, CodDepartamento From Reportes " & _
                                                 "WHERE  (KeyGrupo = '" & CodigoGrupo & "') AND (CodDepartamento = '" & CodDepartamento & "') ORDER BY Orden"
           MDIPrimero.AdoConsulta.Refresh
           If MDIPrimero.AdoConsulta.Recordset.EOF Then
               Orden = Orden + 1
               FrmReportes.DtaReportes.Refresh
    '           FrmReportes.DtaReportes.Recordset.MoveLast
               FrmReportes.DtaReportes.Recordset.AddNew
                  FrmReportes.DtaReportes.Recordset("Descripcion") = DescripcionDepartamento
                  FrmReportes.DtaReportes.Recordset("KeyGrupo") = CodigoGrupo
                  FrmReportes.DtaReportes.Recordset("Nivel") = Nivel + 1
                  FrmReportes.DtaReportes.Recordset("Orden") = Orden
                  FrmReportes.DtaReportes.Recordset("CodDepartamento") = CodDepartamento
                  If Not KeySuperior = "" Then
                    FrmReportes.DtaReportes.Recordset("KeyGrupoSuperior") = KeySuperior
                  End If
                FrmReportes.DtaReportes.Recordset.Update
            Else
            
            
            
            
            End If
           
           
       


       If Not IsNull(FrmReportes.DtaConsulta.Recordset("UbicacionReporte")) Then
        UbicacionReporte = FrmReportes.DtaConsulta.Recordset("UbicacionReporte")
       Else
          Select Case FrmReportes.DtaConsulta.Recordset("TipoCuenta")
            Case "Caja": UbicacionReporte = "Cajas"
            Case "Bancos": UbicacionReporte = "Bancos"
            Case "Cuentas x Cobrar": UbicacionReporte = "Cuentas x Cobrar"
            Case "Inventario": UbicacionReporte = "Inventario"
            Case "Activo Fijo": UbicacionReporte = "Terreno y Edificios"
            Case "Papeleria - Utiles": UbicacionReporte = "Papeleria y Utiles de Oficina"
            Case "Otros Activos": UbicacionReporte = "Otros Activos"
            Case "Cuentas x Pagar": UbicacionReporte = "Proveedores"
            Case "Pasivo": UbicacionReporte = "Pasivos Acumulados"
            Case "Otros Pasivos": UbicacionReporte = "Otros Pasivos"
            Case "Capital": UbicacionReporte = "Acciones Comunes"
            Case "Ingresos - Ventas": UbicacionReporte = "Ingresos - Ventas"
            Case "Costos": UbicacionReporte = "Costos"
            Case "Gastos": UbicacionReporte = "Gastos"
            
        End Select
       End If
       
       If UbicacionReporte = " " Then
         Select Case FrmReportes.DtaConsulta.Recordset("TipoCuenta")
            Case "Caja": UbicacionReporte = "Cajas"
            Case "Bancos": UbicacionReporte = "Bancos"
            Case "Cuentas x Cobrar": UbicacionReporte = "Cuentas x Cobrar"
            Case "Inventario": UbicacionReporte = "Inventario"
            Case "Activo Fijo": UbicacionReporte = "Terreno y Edificios"
            Case "Papeleria - Utiles": UbicacionReporte = "Papeleria y Utiles de Oficina"
            Case "Otros Activos": UbicacionReporte = "Otros Activos"
            Case "Cuentas x Pagar": UbicacionReporte = "Proveedores"
            Case "Pasivo": UbicacionReporte = "Pasivos Acumulados"
            Case "Otros Pasivos": UbicacionReporte = "Otros Pasivos"
            Case "Capital": UbicacionReporte = "Acciones Comunes"
            Case "Ingresos - Ventas": UbicacionReporte = "Ingresos - Ventas"
            Case "Costos": UbicacionReporte = "Costos"
            Case "Gastos": UbicacionReporte = "Gastos"
            
        End Select
       ElseIf UbicacionReporte = "" Then
         Select Case FrmReportes.DtaConsulta.Recordset("TipoCuenta")
            Case "Caja": UbicacionReporte = "Cajas"
            Case "Bancos": UbicacionReporte = "Bancos"
            Case "Cuentas x Cobrar": UbicacionReporte = "Cuentas x Cobrar"
            Case "Inventario": UbicacionReporte = "Inventario"
            Case "Activo Fijo": UbicacionReporte = "Terreno y Edificios"
            Case "Papeleria - Utiles": UbicacionReporte = "Papeleria y Utiles de Oficina"
            Case "Otros Activos": UbicacionReporte = "Otros Activos"
            Case "Cuentas x Pagar": UbicacionReporte = "Proveedores"
            Case "Pasivo": UbicacionReporte = "Pasivos Acumulados"
            Case "Otros Pasivos": UbicacionReporte = "Otros Pasivos"
            Case "Capital": UbicacionReporte = "Acciones Comunes"
            Case "Ingresos - Ventas": UbicacionReporte = "Ingresos - Ventas"
            Case "Costos": UbicacionReporte = "Costos"
            Case "Gastos": UbicacionReporte = "Gastos"
            
        End Select
      End If

       
'/////////////Lleno de Espacios el codigo de la cuenta//////////////////////////////
           For i = 1 To Nivel
             DescripCuenta = " " + DescripCuenta
           Next
           
            Orden = Orden + 1
           FrmReportes.DtaReportes.Recordset.AddNew
              FrmReportes.DtaReportes.Recordset("Descripcion") = DescripCuenta
              FrmReportes.DtaReportes.Recordset("KeyGrupo") = CodigoCuenta
              FrmReportes.DtaReportes.Recordset("KeyGrupoCuenta") = CodigoGrupo
              FrmReportes.DtaReportes.Recordset("Nivel") = Nivel + 1
              FrmReportes.DtaReportes.Recordset("Ubicacion") = UbicacionReporte
              FrmReportes.DtaReportes.Recordset("Orden") = Orden
              FrmReportes.DtaReportes.Recordset("CodCuentas") = CodigoCuenta
              FrmReportes.DtaReportes.Recordset("CodDepartamento") = CodDepartamento
            FrmReportes.DtaReportes.Recordset.Update
            
           FrmReportes.DtaConsulta.Recordset.MoveNext
           
           
            '-------------------------VERIFICO SI EL REGISTRO SIGUIENTE ---------------------------
           '-------------------------CAMBIO A OTRO DEPARTAMETO ------------------------------
           '---------------------------------------------------------------------------------------
           If FrmReportes.DtaConsulta.Recordset.EOF Then

            
           '------------------------------------------------------------------------------
           '--------------------------------------SI ES EL FIN DEL GRUPO TOTALIZO DEPTARMANTOS --
           '------------------------------------------------------------------------------------
           
           
              MDIPrimero.AdoConsulta.RecordSource = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, Nivel, CodCuentas, CodDepartamento From Reportes " & _
                                                    "WHERE  (KeyGrupo = '" & CodigoGrupo & "') AND (CodDepartamento = '" & CodDepartamento & "') AND  (Descripcion LIKE '%Total%') ORDER BY Orden"
              MDIPrimero.AdoConsulta.Refresh
              If MDIPrimero.AdoConsulta.Recordset.EOF Then
                  Orden = Orden + 1
                  FrmReportes.DtaReportes.Refresh
    '             FrmReportes.DtaReportes.Recordset.MoveLast
                  FrmReportes.DtaReportes.Recordset.AddNew
                    FrmReportes.DtaReportes.Recordset("Descripcion") = TotalDescripcionDpto
                    FrmReportes.DtaReportes.Recordset("KeyGrupo") = CodigoGrupo
                    FrmReportes.DtaReportes.Recordset("Nivel") = Nivel + 1
                    FrmReportes.DtaReportes.Recordset("Orden") = Orden
                    FrmReportes.DtaReportes.Recordset("CodDepartamento") = CodDepartamento
                    If Not KeySuperior = "" Then
                      FrmReportes.DtaReportes.Recordset("KeyGrupoSuperior") = KeySuperior
                    End If
                  FrmReportes.DtaReportes.Recordset.Update
              End If

           Else
           
               CodigoGrupoAnt = CodigoGrupo
               CodDepartamentoAnt = CodDepartamento
               If Not IsNull(FrmReportes.DtaConsulta.Recordset("VoucherNo")) Then
                 CodDepartamento = FrmReportes.DtaConsulta.Recordset("VoucherNo")
               Else
                 CodDepartamento = "-"
               End If
               If CodDepartamento <> CodDepartamentoAnt Then
                  MDIPrimero.AdoConsulta.RecordSource = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, Nivel, CodCuentas, CodDepartamento From Reportes " & _
                                                        "WHERE  (KeyGrupo = '" & CodigoGrupoAnt & "') AND (CodDepartamento = '" & CodDepartamentoAnt & "') AND  (Descripcion LIKE '%Total%') ORDER BY Orden"
                  MDIPrimero.AdoConsulta.Refresh
                  If MDIPrimero.AdoConsulta.Recordset.EOF Then
                      Orden = Orden + 1
                      FrmReportes.DtaReportes.Refresh
        '             FrmReportes.DtaReportes.Recordset.MoveLast
                      FrmReportes.DtaReportes.Recordset.AddNew
                        FrmReportes.DtaReportes.Recordset("Descripcion") = TotalDescripcionDpto
                        FrmReportes.DtaReportes.Recordset("KeyGrupo") = CodigoGrupoAnt
                        FrmReportes.DtaReportes.Recordset("Nivel") = Nivel + 1
                        FrmReportes.DtaReportes.Recordset("Orden") = Orden
                        FrmReportes.DtaReportes.Recordset("CodDepartamento") = CodDepartamentoAnt
                        If Not KeySuperior = "" Then
                          FrmReportes.DtaReportes.Recordset("KeyGrupoSuperior") = KeySuperior
                        End If
                      FrmReportes.DtaReportes.Recordset.Update
                  End If
               
               
               End If
           End If
           
           
           
           
           
       Loop
  

'///////////Busco si este Grupo no tiene SubGrupos para totalizar/////////////////////////////////////////////
       FrmReportes.DtaConsulta.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.DescripcionGrupo From Grupos Where (((Grupos.KeyGrupoSuperior) = '" & CodigoGrupo & "'))"
       FrmReportes.DtaConsulta.Refresh
       If FrmReportes.DtaConsulta.Recordset.EOF Then
'/////////////////////////////Verifico si es de Balance, para Agregar Capital/////
          If QUIEN = "Balance" Then
            If CodigoGrupo = "C" Then
             DescripcionBalance = " Resultado Periodo"

               DescripcionBalance = " " + DescripcionBalance
               
              Orden = Orden + 1
              FrmReportes.DtaReportes.Recordset.AddNew
               FrmReportes.DtaReportes.Recordset("Descripcion") = DescripcionBalance
               FrmReportes.DtaReportes.Recordset("KeyGrupo") = CodigoGrupo
               FrmReportes.DtaReportes.Recordset("Nivel") = Nivel
              FrmReportes.DtaReportes.Recordset("Orden") = Orden
              FrmReportes.DtaReportes.Recordset("Ubicacion") = "Resultado Periodo"
              FrmReportes.DtaReportes.Recordset.Update
              
 
 
            End If
          End If
         
          
          
          
            Orden = Orden + 1
            FrmReportes.DtaReportes.Refresh
            FrmReportes.DtaReportes.Recordset.AddNew
              FrmReportes.DtaReportes.Recordset("Descripcion") = TotalDescripcion
              FrmReportes.DtaReportes.Recordset("KeyGrupo") = CodigoGrupo
              FrmReportes.DtaReportes.Recordset("Nivel") = Nivel
              FrmReportes.DtaReportes.Recordset("Orden") = Orden
              If Not KeySuperior = "" Then
               FrmReportes.DtaReportes.Recordset("KeyGrupoSuperior") = KeySuperior
              End If
            FrmReportes.DtaReportes.Recordset.Update
       
       End If
       
'///////////Busco si el padre de este grupo ya totalizaron todos sus hijos/////////////////////////////////////////////
       FrmReportes.DtaReportes.Refresh
       FrmReportes.DtaConsulta.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.DescripcionGrupo From Grupos Where (((Grupos.KeyGrupoSuperior) = '" & KeySuperior & "')) ORDER BY KeyGrupo"
       FrmReportes.DtaConsulta.Refresh
       If Not FrmReportes.DtaConsulta.Recordset.EOF Then
     
         FrmReportes.DtaConsulta.Recordset.MoveLast
         NumeroHijos = FrmReportes.DtaConsulta.Recordset.RecordCount
         FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.KeyGrupo From Reportes Where (Reportes.KeyGrupoSuperior = '" & KeySuperior & "') AND (CodDepartamento IS NULL)"
         FrmReportes.DtaConsulta2.Refresh
         If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
          FrmReportes.DtaConsulta2.Recordset.MoveLast
          NumeroHijosTotales = Val(FrmReportes.DtaConsulta2.Recordset.RecordCount) / 2
         End If
         
'/////////////////////Verifico todos los hijos ya tienen totales////////////////////////////////////
       If NumeroHijos = NumeroHijosTotales Then
            FrmReportes.DtaGrupos.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.DescripcionGrupo From Grupos Where (((Grupos.KeyGrupo) = '" & KeySuperior & "'))"
            FrmReportes.DtaGrupos.Refresh
            If Not FrmReportes.DtaGrupos.Recordset.EOF Then
              DescripcionPadre = "Total " + FrmReportes.DtaGrupos.Recordset("DescripcionGrupo")
'//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
              Longitud = Len(FrmReportes.DtaGrupos.Recordset("KeyGrupo"))
              If Longitud > 1 Then
                 If Longitud = 5 Then
                     Nivel = 2
                 Else
                     Nivel = (Longitud - 5) / 2
                     Nivel = Nivel + 2
                 End If
               Else
                    Nivel = 1
               End If
'////////////////////Lleno de Espacios la Descripcion del padre///////////////////////////////////
             For i = 2 To Nivel
               DescripcionPadre = " " + DescripcionPadre
             Next

'/////////////////////////////Verifico si es de Balance, para Agregar Capital/////
              If QUIEN = "Balance" Then
                If KeySuperior = "C" Then
                 DescripcionBalance = "Resultado Periodo"
    
                   DescripcionBalance = " " + DescripcionBalance
                  FrmReportes.DtaReportes.Refresh
                  Orden = Orden + 1
                  FrmReportes.DtaReportes.Recordset.AddNew
                   FrmReportes.DtaReportes.Recordset("Descripcion") = DescripcionBalance
                   FrmReportes.DtaReportes.Recordset("KeyGrupo") = KeySuperior
                  FrmReportes.DtaReportes.Recordset("Nivel") = Nivel
                  FrmReportes.DtaReportes.Recordset("Orden") = Orden
                  FrmReportes.DtaReportes.Recordset("Ubicacion") = "Resultado Periodo"
                  FrmReportes.DtaReportes.Recordset.Update
                  
    
     
                End If
              End If
'/////////////////////pongo el total del padre////////////////////////////
                Orden = Orden + 1
                FrmReportes.DtaReportes.Recordset.AddNew
                 FrmReportes.DtaReportes.Recordset("Descripcion") = DescripcionPadre
                 FrmReportes.DtaReportes.Recordset("KeyGrupo") = KeySuperior
                FrmReportes.DtaReportes.Recordset("Nivel") = Nivel
                FrmReportes.DtaReportes.Recordset("Orden") = Orden
                 If Not IsNull(FrmReportes.DtaGrupos.Recordset("KeyGrupoSuperior")) Then
                  FrmReportes.DtaReportes.Recordset("KeyGrupoSuperior") = FrmReportes.DtaGrupos.Recordset("KeyGrupoSuperior")
                 End If
                FrmReportes.DtaReportes.Recordset.Update
                FrmReportes.DtaReportes.Refresh
 
 
 '/////////////////////////Busco Si ya totalizaron los otros padres//////////////////////
                If Not IsNull(FrmReportes.DtaGrupos.Recordset("KeyGrupoSuperior")) Then
                   Respuesta = FrmReportes.DtaGrupos.Recordset("KeyGrupoSuperior")
                End If
       
       KeyUltimo = KeySuperior
       Nivel = Nivel - 1
       For i = Nivel To 1 Step -1
'///////Verifico si el Grupo Anterior es el ultimo Hijo del Padre acutal

        FrmReportes.DtaConsulta.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.DescripcionGrupo From Grupos Where (((Grupos.KeyGrupoSuperior) = '" & Respuesta & "'))"
        FrmReportes.DtaConsulta.Refresh
        If Not FrmReportes.DtaConsulta.Recordset.EOF Then
           FrmReportes.DtaConsulta.Recordset.MoveLast
           KeyHijo = FrmReportes.DtaConsulta.Recordset("KeyGrupo")
           If KeyHijo = KeyUltimo Then
'////////Si es el ultimo Hijo Busco los datos del padre para Cerrarlo////////////////////////////
              KeySuperior = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
              FrmReportes.DtaConsulta2.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.DescripcionGrupo From Grupos Where (((Grupos.KeyGrupo) = '" & KeySuperior & "'))"
              FrmReportes.DtaConsulta2.Refresh
              DescripcionPadre = "Total " + FrmReportes.DtaConsulta2.Recordset("DescripcionGrupo")

'////////////////////Lleno de Espacios la Descripcion del padre///////////////////////////////////
             For J = 2 To i
               DescripcionPadre = " " + DescripcionPadre
             Next
           
'/////////////////////////////Verifico si es de Balance, para Agregar Capital/////
          If QUIEN = "Balance" Then
            If KeySuperior = "C" Then
             DescripcionBalance = "Resultado Periodo"

               DescripcionBalance = " " + DescripcionBalance
              Orden = Orden + 1
              FrmReportes.DtaReportes.Recordset.AddNew
               FrmReportes.DtaReportes.Recordset("Descripcion") = DescripcionBalance
               FrmReportes.DtaReportes.Recordset("KeyGrupo") = "RP"
              FrmReportes.DtaReportes.Recordset("Nivel") = Nivel
              FrmReportes.DtaReportes.Recordset("Orden") = Orden
              FrmReportes.DtaReportes.Recordset("Ubicacion") = "Resultado Periodo"
              FrmReportes.DtaReportes.Recordset.Update
         
            End If
          End If
'/////////////////////pongo el total del padre////////////////////////////
               Orden = Orden + 1
               FrmReportes.DtaReportes.Recordset.AddNew
               FrmReportes.DtaReportes.Recordset("Descripcion") = DescripcionPadre
               FrmReportes.DtaReportes.Recordset("KeyGrupo") = KeySuperior
              FrmReportes.DtaReportes.Recordset("Nivel") = i
              FrmReportes.DtaReportes.Recordset("Orden") = Orden
               If Not IsNull(FrmReportes.DtaConsulta2.Recordset("KeyGrupoSuperior")) Then
                FrmReportes.DtaReportes.Recordset("KeyGrupoSuperior") = FrmReportes.DtaConsulta2.Recordset("KeyGrupoSuperior")
               End If
              FrmReportes.DtaReportes.Recordset.Update
              KeyUltimo = KeySuperior
              FrmReportes.DtaReportes.Refresh
           
           Else
            Exit For
           End If
        
             

           If Not IsNull(FrmReportes.DtaConsulta2.Recordset("KeyGrupoSuperior")) Then
              Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupoSuperior")
           End If
        End If
       Next
       
       

            End If
         End If
       End If
       
       
       
       
       
       FrmReportes.DtaHistorial.Recordset.MoveNext
     Loop




End Sub
Public Sub SaldoReportes(QUIEN As String)
Dim CodigoGrupo As String, Sql As String, Fechas As Date
Dim Nivel As Integer, Longitud As Integer, Fecha1 As String
Dim TotalMayor() As String, TotalDescripcion As String
Dim KeySuperior As String, NumeroHijos As Double, NumeroHijosTotales As Double
Dim DescripCuenta As String, DescripcionPadre As String, KeyUltimo As String, CodigoCuentaDesde As String, CodigoCuentaHasta As String
Dim DebitoD As Double, CreditoD As Double, Ajuste As String


  '   ////////////////Elimino los registros del reporte///////////////////
  'frmreportes.DtaElimina.RecordSource = "DELETE Reportes.* From Reportes"
  'frmreportes.DtaElimina.Recordset.Updatable
  
    Dim Orden As Integer  'sirve para ordenar las cuentas
                Orden = 1
             
                NumFecha1 = FechaIni
                NumFecha2 = FechaFin
                
                         If FrmReportes.CmbMoneda.Text = "Córdobas" Then
                           Ajuste = "Dólares"
                         ElseIf FrmReportes.CmbMoneda.Text = "Dólares" Then
                            Ajuste = "Córdobas"
                         End If
 
                'Busco que cuentas tienen saldo
                If QUIEN = "Balance" Then
                 Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda " & vbLf
                 Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Otros Activos' Or (Cuentas.TipoCuenta) = 'Caja' Or (Cuentas.TipoCuenta) = 'Bancos' Or (Cuentas.TipoCuenta) = 'Cuentas x Cobrar' Or (Cuentas.TipoCuenta) = 'Inventario' Or (Cuentas.TipoCuenta) = 'Papeleria - Utiles' Or (Cuentas.TipoCuenta) = 'Activo Fijo' Or (Cuentas.TipoCuenta) = 'Otros Pasivos' Or (Cuentas.TipoCuenta) = 'Cuentas x Pagar' Or (Cuentas.TipoCuenta) = 'Pasivo' Or (Cuentas.TipoCuenta) = 'Capital')) ORDER BY Cuentas.CodCuentas"
                 FrmReportes.DtaHistorial.RecordSource = Sql
                 FrmReportes.DtaHistorial.Refresh
                ElseIf QUIEN = "Utilidad" Then
                 Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda " & vbLf
                 Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Ingresos - Ventas' Or (Cuentas.TipoCuenta) = 'Costos' Or (Cuentas.TipoCuenta) = 'Gastos')) ORDER BY Cuentas.CodCuentas"
                 FrmReportes.DtaHistorial.RecordSource = Sql
                 FrmReportes.DtaHistorial.Refresh
                ElseIf QUIEN = "Resultado" Then
                 Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda " & vbLf
                 Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Ingresos - Ventas' Or (Cuentas.TipoCuenta) = 'Costos' Or (Cuentas.TipoCuenta) = 'Gastos')) ORDER BY Cuentas.CodCuentas"
                 FrmReportes.DtaHistorial.RecordSource = Sql
                 FrmReportes.DtaHistorial.Refresh
                ElseIf QUIEN = "UtilidadResultado" Then
                 Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda " & vbLf
                 Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Ingresos - Ventas' Or (Cuentas.TipoCuenta) = 'Costos' Or (Cuentas.TipoCuenta) = 'Gastos')) ORDER BY Cuentas.CodCuentas"
                 FrmReportes.DtaHistorial.RecordSource = Sql
                 FrmReportes.DtaHistorial.Refresh
                ElseIf QUIEN = "ResultadoDpto" Then
                 Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Transacciones.VoucherNo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Transacciones.VoucherNo " & vbLf
                 Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Ingresos - Ventas' Or (Cuentas.TipoCuenta) = 'Costos' Or (Cuentas.TipoCuenta) = 'Gastos')) ORDER BY Cuentas.CodCuentas, Transacciones.VoucherNo"
                 FrmReportes.DtaHistorial.RecordSource = Sql
                 FrmReportes.DtaHistorial.Refresh
                 
                 
                 
                 
                ElseIf QUIEN = "Balanza" Then
                
                     If FrmReportes.TxtDesde.Text = "" Then
                       FrmReportes.DtaConsulta.RecordSource = "SELECT * From Grupos ORDER BY KeyGrupo"
                       FrmReportes.DtaConsulta.Refresh
                       If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                         FrmReportes.DtaConsulta.Recordset.MoveFirst
                         CodigoCuentaDesde = FrmReportes.DtaConsulta.Recordset("KeyGrupo")
                       End If
                    Else
                        CodigoCuentaDesde = FrmReportes.TxtKeyGrupoDesde.Text
                    End If
                       
                    If FrmReportes.TxtHasta.Text = "" Then
                       FrmReportes.DtaConsulta.RecordSource = "SELECT * From Grupos ORDER BY KeyGrupo"
                       FrmReportes.DtaConsulta.Refresh
                       If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                         FrmReportes.DtaConsulta.Recordset.MoveLast
                         CodigoCuentaHasta = FrmReportes.DtaConsulta.Recordset("KeyGrupo")
                       End If
                    Else
                       CodigoCuentaHasta = FrmReportes.TxtKeyGrupoHasta.Text
                    End If
                
                
                
                    FrmReportes.DtaHistorial.RecordSource = "SELECT  Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 5)) AS MDebito, SUM(ROUND(Transacciones.Credito * Transacciones.TCambio,5)) AS MCredito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 5) - ROUND(Transacciones.Credito * Transacciones.TCambio, 5)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.KeyGrupo) As KeyGrupo FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
                                                            "WHERE  (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') GROUP BY Cuentas.CodCuentas HAVING  (MAX(Cuentas.KeyGrupo) BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') ORDER BY Cuentas.CodCuentas"
                    FrmReportes.DtaHistorial.Refresh
                 
                    If FrmReportes.CmbMoneda.Text = "Córdobas" Then
                        If FrmReportes.ChkQuitarMovimiento.Value = 1 Then

                          ConsultaTotalesMovimientos = "SELECT Transacciones.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 3)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito, 3)) AS MCredito, MAX(IndiceTransaccion.Fuente) AS Fuente, MAX(Cuentas.KeyGrupo) AS KeyGrupo FROM  Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING (MAX(IndiceTransaccion.Fuente) <> 'Cierre') AND (MAX(Cuentas.KeyGrupo) BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') ORDER BY Transacciones.CodCuentas"
                        Else
                          ConsultaTotalesMovimientos = "SELECT Transacciones.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 3)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito, 3)) AS MCredito, MAX(Cuentas.KeyGrupo) AS KeyGrupo FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING (MAX(Cuentas.KeyGrupo) BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "')"
                        End If
'                   Else
                        ConsultaTotalesMovimientos = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas,5)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas,5)) AS MCredito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas,5) - ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas,5)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN  Tasas ON Transacciones.FechaTransaccion = Tasas.FechaTasas  " & _
                                                     "WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyy-mm-dd") & "' AND '" & Format(FechaFin, "yyyy-mm-dd") & "') GROUP BY Cuentas.CodCuentas"
                          
                    End If
                    
                    
              ElseIf QUIEN = "BalanzaLibroMayor" Then
                
                
                
                     If FrmReportes.TxtDesde.Text = "" Then
                       FrmReportes.DtaConsulta.RecordSource = "SELECT * From Grupos ORDER BY KeyGrupo"
                       FrmReportes.DtaConsulta.Refresh
                       If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                         FrmReportes.DtaConsulta.Recordset.MoveFirst
                         CodigoCuentaDesde = FrmReportes.DtaConsulta.Recordset("KeyGrupo")
                       End If
                    Else
                        CodigoCuentaDesde = FrmReportes.TxtKeyGrupoDesde.Text
                    End If
                       
                    If FrmReportes.TxtHasta.Text = "" Then
                       FrmReportes.DtaConsulta.RecordSource = "SELECT * From Grupos ORDER BY KeyGrupo"
                       FrmReportes.DtaConsulta.Refresh
                       If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                         FrmReportes.DtaConsulta.Recordset.MoveLast
                         CodigoCuentaHasta = FrmReportes.DtaConsulta.Recordset("KeyGrupo")
                       End If
                    Else
                       CodigoCuentaHasta = FrmReportes.TxtKeyGrupoHasta.Text
                    End If
                
                
                
                    FrmReportes.DtaHistorial.RecordSource = "SELECT  Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 5)) AS MDebito, SUM(ROUND(Transacciones.Credito * Transacciones.TCambio,5)) AS MCredito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 5) - ROUND(Transacciones.Credito * Transacciones.TCambio, 5)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.KeyGrupo) As KeyGrupo FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
                                                            "WHERE  (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') GROUP BY Cuentas.CodCuentas HAVING  (MAX(Cuentas.KeyGrupo) BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') ORDER BY Cuentas.CodCuentas"
                    FrmReportes.DtaHistorial.Refresh
                 
                    If FrmReportes.CmbMoneda.Text = "Córdobas" Then
                        If FrmReportes.ChkQuitarMovimiento.Value = 1 Then

                          ConsultaTotalesMovimientos = "SELECT Transacciones.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 3)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito, 3)) AS MCredito, MAX(IndiceTransaccion.Fuente) AS Fuente, MAX(Cuentas.KeyGrupo) AS KeyGrupo FROM  Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas, Transacciones.NPeriodo HAVING (MAX(IndiceTransaccion.Fuente) <> 'Cierre') AND (MAX(Cuentas.KeyGrupo) BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') ORDER BY Transacciones.CodCuentas"
                        Else
                          ConsultaTotalesMovimientos = "SELECT Transacciones.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 3)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito, 3)) AS MCredito, MAX(Cuentas.KeyGrupo) AS KeyGrupo FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Transacciones.CodCuentas, Transacciones.NPeriodo HAVING (MAX(Cuentas.KeyGrupo) BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "')"
                        End If
'                   Else
                        ConsultaTotalesMovimientos = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas,5)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas,5)) AS MCredito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas,5) - ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas,5)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN  Tasas ON Transacciones.FechaTransaccion = Tasas.FechaTasas  " & _
                                                     "WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyy-mm-dd") & "' AND '" & Format(FechaFin, "yyyy-mm-dd") & "') GROUP BY Cuentas.CodCuentas, Transacciones.NPeriodo"
                          
                    End If
                 
                ElseIf QUIEN = "BalanzaCodigo" Then
                
                            If FrmReportes.DBCodigo.Text = "" Then
                                 FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.* From Cuentas ORDER BY CodCuentas"
                                 FrmReportes.DtaConsulta.Refresh
                                 If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                                   FrmReportes.DtaConsulta.Recordset.MoveFirst
                                   CodigoCuentaDesde = FrmReportes.DtaConsulta.Recordset("CodCuentas")
                                End If
                            Else
                              CodigoCuentaDesde = FrmReportes.DBCodigo.Text
                            End If
                            
                            If FrmReportes.DBCodigoHasta.Text = "" Then
                                 FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.* From Cuentas ORDER BY CodCuentas"
                                 FrmReportes.DtaConsulta.Refresh
                                 If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                                   FrmReportes.DtaConsulta.Recordset.MoveLast
                                   CodigoCuentaHasta = FrmReportes.DtaConsulta.Recordset("CodCuentas")
                                End If
                            Else
                               CodigoCuentaHasta = FrmReportes.DBCodigoHasta.Text
                            End If
                            
                              
                             FrmReportes.DtaHistorial.RecordSource = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 5)) AS MDebito, SUM(ROUND(Transacciones.Credito * Transacciones.TCambio,5)) AS MCredito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 5) - ROUND(Transacciones.Credito * Transacciones.TCambio, 5)) AS Total,MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Cuentas.TipoCuenta) AS TipoCuenta FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') GROUP BY Cuentas.CodCuentas HAVING (Cuentas.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') ORDER BY Cuentas.CodCuentas"
                             FrmReportes.DtaHistorial.Refresh
                                '///////////////////////////////////////////////////////////////////////////////
                                '////////////////guardo la consulta para actualizar en el reporte///////////////
                                '///////////////////////////////////////////////////////////////////////////////
                                
                                If FrmReportes.CmbMoneda.Text = "Córdobas" Then
                                    If FrmReportes.ChkQuitarMovimiento.Value = 1 Then
                                      ConsultaTotalesMovimientos = "SELECT Transacciones.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 3)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito, 3)) AS MCredito, MAX(IndiceTransaccion.Fuente) AS Fuente FROM  Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Transacciones.CodCuentas, IndiceTransaccion.Ajuste HAVING (MAX(IndiceTransaccion.Fuente) <> 'Cierre') AND (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') AND (IndiceTransaccion.Ajuste <> '" & Ajuste & "') ORDER BY Transacciones.CodCuentas"
                                    Else
                                      ConsultaTotalesMovimientos = "SELECT Transacciones.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 3)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito, 3)) AS MCredito FROM  Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento  WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Transacciones.CodCuentas, IndiceTransaccion.Ajuste HAVING (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') AND (IndiceTransaccion.Ajuste <> '" & Ajuste & "')"
                                    End If
                                Else
                                     ConsultaTotalesMovimientos = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 5)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas, 5)) AS MCredito,SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 5) - ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas, 5)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTransaccion = Tasas.FechaTasas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo " & _
                                                                 "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Cuentas.CodCuentas, IndiceTransaccion.Ajuste HAVING (Cuentas.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') AND (IndiceTransaccion.Ajuste <> '" & Ajuste & "')"
                                End If
                
                
                ElseIf QUIEN = "SaldoCuentas" Then
                            QUIEN = "BalanzaCodigo"
                            If FrmReportes.DBCodigo.Text = "" Then
                                 FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.* From Cuentas ORDER BY CodCuentas"
                                 FrmReportes.DtaConsulta.Refresh
                                 If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                                   FrmReportes.DtaConsulta.Recordset.MoveFirst
                                   CodigoCuentaDesde = FrmReportes.DtaConsulta.Recordset("CodCuentas")
                                End If
                            Else
                              CodigoCuentaDesde = FrmReportes.DBCodigo.Text
                            End If
                            
                            If FrmReportes.DBCodigoHasta.Text = "" Then
                                 FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.* From Cuentas ORDER BY CodCuentas"
                                 FrmReportes.DtaConsulta.Refresh
                                 If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                                   FrmReportes.DtaConsulta.Recordset.MoveLast
                                   CodigoCuentaHasta = FrmReportes.DtaConsulta.Recordset("CodCuentas")
                                End If
                            Else
                               CodigoCuentaHasta = FrmReportes.DBCodigoHasta.Text
                            End If
                            
                              
                             FrmReportes.DtaHistorial.RecordSource = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 5)) AS MDebito, SUM(ROUND(Transacciones.Credito * Transacciones.TCambio,5)) AS MCredito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 5) - ROUND(Transacciones.Credito * Transacciones.TCambio, 5)) AS Total,MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Cuentas.TipoCuenta) AS TipoCuenta FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') GROUP BY Cuentas.CodCuentas HAVING (Cuentas.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "')  ORDER BY Cuentas.CodCuentas"  'AND (MAX(Cuentas.TipoCuenta) = 'Cuentas x Cobrar')
                             FrmReportes.DtaHistorial.Refresh
                                '///////////////////////////////////////////////////////////////////////////////
                                '////////////////guardo la consulta para actualizar en el reporte///////////////
                                '///////////////////////////////////////////////////////////////////////////////
                                
                                If FrmReportes.CmbMoneda.Text = "Córdobas" Then
                                    If FrmReportes.ChkQuitarMovimiento.Value = 1 Then
                                      ConsultaTotalesMovimientos = "SELECT Transacciones.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 3)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito, 3)) AS MCredito, MAX(IndiceTransaccion.Fuente) AS Fuente FROM  Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING (MAX(IndiceTransaccion.Fuente) <> 'Cierre') AND (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') ORDER BY Transacciones.CodCuentas"
                                    Else
                                      ConsultaTotalesMovimientos = "SELECT CodCuentas, SUM(ROUND(Debito * TCambio, 3)) AS MDebito, SUM(ROUND(TCambio * Credito, 3)) AS MCredito From Transacciones  WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY CodCuentas HAVING (CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') "
                                    End If
                                Else
                                     ConsultaTotalesMovimientos = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 5)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas, 5)) AS MCredito,SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 5) - ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas, 5)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTransaccion = Tasas.FechaTasas " & _
                                                                 "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Cuentas.CodCuentas HAVING (Cuentas.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') "
                                End If
                
                
                
                Else
                     FrmReportes.DtaHistorial.RecordSource = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio,5)) AS MDebito, SUM(ROUND(Transacciones.Credito * Transacciones.TCambio,5)) AS MCredito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio,5) - ROUND(Transacciones.Credito * Transacciones.TCambio,5)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Cuentas.TipoCuenta) AS TipoCuenta FROM  Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
                                                             "WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') GROUP BY Cuentas.CodCuentas ORDER BY Cuentas.CodCuentas"
                     FrmReportes.DtaHistorial.Refresh
                        '///////////////////////////////////////////////////////////////////////////////
                        '////////////////guardo la consulta para actualizar en el reporte///////////////
                        '///////////////////////////////////////////////////////////////////////////////
                        
                        If FrmReportes.CmbMoneda.Text = "Córdobas" Then
    '                         ConsultaTotalesMovimientos = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio,5)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito,5)) AS MCredito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio,5) - ROUND(Transacciones.TCambio * Transacciones.Credito,5)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda FROM Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
    '                                                      "WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') GROUP BY Cuentas.CodCuentas"
                            If FrmReportes.ChkQuitarMovimiento.Value = 1 Then
                              ConsultaTotalesMovimientos = "SELECT  Transacciones.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 3)) AS MDebito,  SUM(ROUND(Transacciones.TCambio * Transacciones.Credito, 3)) AS MCredito, MAX(IndiceTransaccion.Fuente) AS Fuente FROM  Transacciones INNER JOIN  IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento  " & _
                                                           "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING (MAX(IndiceTransaccion.Fuente) <> 'Cierre') ORDER BY Transacciones.CodCuentas"
                            Else
                              ConsultaTotalesMovimientos = "SELECT CodCuentas, SUM(ROUND(Debito * TCambio, 3)) AS MDebito, SUM(ROUND(TCambio * Credito, 3)) AS MCredito From Transacciones  WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY CodCuentas"
                            End If
    '                          ConsultaTotalesMovimientos = "SELECT   NTransaccion, FechaTransaccion, VoucherNo, ChequeNo, DescripcionMovimiento, CodCuentas, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, Debito - Credito AS Balance, TCambio, NumeroMovimiento, Beneficiario From Transacciones WHERE (FechaTransaccion BETWEEN '" & Format(FechaFin, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') ORDER BY CodCuentas, FechaTransaccion, NTransaccion"
                        Else
                            ConsultaTotalesMovimientos = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas,5)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas,5)) AS MCredito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas,5) - ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas,5)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN  Tasas ON Transacciones.FechaTransaccion = Tasas.FechaTasas  " & _
                                                         "WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyy-mm-dd") & "' AND '" & Format(FechaFin, "yyyy-mm-dd") & "') GROUP BY Cuentas.CodCuentas"
                              
                        End If
            
                End If

               Totalingresos = 0
               TotalGastos = 0
               FrmReportes.osProgress1.Value = 0
               FrmReportes.osProgress1.Visible = True
            
               If Not FrmReportes.DtaHistorial.Recordset.EOF Then
                FrmReportes.DtaHistorial.Recordset.MoveLast
                FrmReportes.osProgress1.Max = FrmReportes.DtaHistorial.Recordset.RecordCount
                FrmReportes.DtaHistorial.Recordset.MoveFirst
               End If
               
'*************************************************************************************************************
'*************************************************************************************************************
'*************************************************************************************************************
        '/////////////////////////////////////////////////////////////////////////////////////////////////////
        '//////////////////MOVIMIENTOS DEL PERIODO////////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////////////////////
'*************************************************************************************************************
'*************************************************************************************************************
'*************************************************************************************************************
 
         Do While Not FrmReportes.DtaHistorial.Recordset.EOF
                   NumFecha1 = FechaIni
                   NumFecha2 = FechaFin
                    '////////Consulto los registros del periodo seleccionado.///////////
                    CodigoCuenta = FrmReportes.DtaHistorial.Recordset("CodCuentas")
                    
                    If CodigoCuenta = "6500" Then
                      CodigoCuenta = "6500"
                    End If
                    
                    
                    
                     FrmReportes.LblProgreso.Caption = "Consultando Registros del periodo seleccionado para la cuenta " & CodigoCuenta
                     FrmReportes.osProgress1.Value = FrmReportes.osProgress1.Value + 1
                     DoEvents
                     If FrmReportes.ChkQuitarMovimiento.Value = 1 Then
'                       FrmReportes.DtaConsulta.RecordSource = "SELECT  Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio - Transacciones.TCambio * Transacciones.Credito) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NTransaccion) AS Transaccion, MAX(IndiceTransaccion.Fuente) AS Fuente  FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN  IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion  " & _
'                                                              "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Cuentas.CodCuentas HAVING  (Cuentas.CodCuentas = '" & CodigoCuenta & "') AND (MAX(IndiceTransaccion.Fuente) <> 'Cierre')"
                                               FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio - Transacciones.TCambio * Transacciones.Credito) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Transacciones.FechaTransaccion)  AS FechaTransaccion, MAX(Transacciones.NTransaccion) AS Transaccion, MAX(IndiceTransaccion.Fuente) AS Fuente FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
                                                              "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Cuentas.CodCuentas, IndiceTransaccion.Ajuste HAVING   (Cuentas.CodCuentas = '" & CodigoCuenta & "') AND (MAX(IndiceTransaccion.Fuente) <> 'Cierre') AND (IndiceTransaccion.Ajuste <> '" & Ajuste & "')"
                     Else
'                       FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio - Transacciones.TCambio * Transacciones.Credito) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NTransaccion) As Transaccion FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
'                                                           "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Cuentas.CodCuentas HAVING (Cuentas.CodCuentas = '" & CodigoCuenta & "')"
                                            
'                       FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio - Transacciones.TCambio * Transacciones.Credito) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Transacciones.FechaTransaccion)  AS FechaTransaccion, MAX(Transacciones.NTransaccion) AS Transaccion FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
'                                                              "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) AND (IndiceTransaccion.Ajuste <> '" & Ajuste & "') GROUP BY Cuentas.CodCuentas HAVING   (Cuentas.CodCuentas = '" & CodigoCuenta & "')"
                     
                        FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio,5)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito,5)) AS MCredito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio,5) - ROUND(Transacciones.TCambio * Transacciones.Credito,5)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NTransaccion) As Transaccion FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento  " & _
                                                           "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Cuentas.CodCuentas,IndiceTransaccion.Ajuste HAVING (Cuentas.CodCuentas = '" & CodigoCuenta & "') AND (IndiceTransaccion.Ajuste <> '" & Ajuste & "')"

                     End If


                    FrmReportes.DtaConsulta.Refresh
                     
                    TotalDebitoH = 0
                    TotalCreditoH = 0
                   If FrmReportes.ChkBalanza.Value = 1 Then
                   
                    '////////////////////////////////////////////////////////////////////////////////////////////
                    '///////CON ESTA CONSULTA BUSCO EL HISTORICO DEL PERIODO////////////////////////////////////////
                    '//////////////////////////////////////////////////////////////////////////////////////////////
                     
                     TotalDebitoH = 0
                     TotalCreditoH = 0
                     If FrmReportes.ChkBalanza.Value = 1 Then
'                         FrmReportes.AdoHistorial.RecordSource = "SELECT Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.Credito * Transacciones.TCambio) AS Total,Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, MAX(Transacciones.NTransaccion) AS Transaccion FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (Transacciones.FechaTransaccion < '" & Format(FechaIni, "yyyymmdd") & "') GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda HAVING (Cuentas.CodCuentas = '" & CodigoCuenta & "') ORDER BY Cuentas.CodCuentas"
                        FrmReportes.AdoHistorial.RecordSource = "SELECT Cuentas.CodCuentas, ROUND(Transacciones.Debito * Transacciones.TCambio,5) AS MDebito, ROUND(Transacciones.TCambio * Transacciones.Credito,5) AS MCredito, ROUND(Transacciones.Debito * Transacciones.TCambio,5) - ROUND(Transacciones.TCambio * Transacciones.Credito,5) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Transacciones.NTransaccion AS Transaccion FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento  " & _
                                                                "WHERE (Transacciones.FechaTransaccion < '" & Format(FechaIni, "yyyy-mm-dd") & "') AND (Cuentas.CodCuentas = '" & CodigoCuenta & "') AND (IndiceTransaccion.Ajuste <> '" & Ajuste & "') ORDER BY Cuentas.CodCuentas"
                        FrmReportes.AdoHistorial.Refresh
                        If Not FrmReportes.AdoHistorial.Recordset.EOF Then
                            If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                                If Not IsNull(FrmReportes.AdoHistorial.Recordset("MDebito")) Then
                                    TotalDebitoH = FrmReportes.AdoHistorial.Recordset("MDebito")
                                End If
                                If Not IsNull(FrmReportes.AdoHistorial.Recordset("MCredito")) Then
                                     TotalCreditoH = FrmReportes.AdoHistorial.Recordset("MCredito")
                                End If
                                            
                                TotalHistorico = TotalDebitoH - TotalCreditoH
                             Else
                               '/////////EN CASO QUE NO SEA CUENTA DE ACITVO/////
                                If Not IsNull(FrmReportes.AdoHistorial.Recordset("MDebito")) Then
                                    TotalDebitoH = FrmReportes.AdoHistorial.Recordset("MDebito")
                                End If
                                If Not IsNull(FrmReportes.AdoHistorial.Recordset("MCredito")) Then
                                     TotalCreditoH = FrmReportes.AdoHistorial.Recordset("MCredito")
                                End If
                                    
                                TotalHistorico = TotalCreditoH - TotalDebitoH
                             End If
                     
                        End If
                      End If
                    End If
                    
                    
                    'encuentra los movimientos que se hicieron de una cuenta entre el rango especificado
                    DoEvents
                    
                    TotalCuenta = 0
                    Total1 = 0
                    FrmReportes.osProgress2.Value = 0
                    Do While Not FrmReportes.DtaConsulta.Recordset.EOF
                      FrmReportes.osProgress2.Visible = True
                        If FrmReportes.osProgress2.Value = 0 Then
                '            FrmReportes.lblProgreso.Caption = "Consultando Registros del periodo seleccionado para la cuenta " & CodigoCuenta
                            FrmReportes.osProgress2.Max = FrmReportes.DtaConsulta.Recordset.RecordCount
                            FrmReportes.osProgress2.Value = FrmReportes.osProgress2.Value + 1
                
                             DoEvents
                        Else
                '            FrmReportes.lblProgreso.Caption = "Consultando Registros del periodo seleccionado para la cuenta " & CodigoCuenta
                            FrmReportes.osProgress2.Value = FrmReportes.osProgress2.Value + 1
                            DoEvents
                        End If
                        
                       TotalDebito = 0
                       TotalCredito = 0
                      TipoCuenta = FrmReportes.DtaConsulta.Recordset("TipoCuenta")
                     
                      TipoMoneda = FrmReportes.DtaConsulta.Recordset("TipoMoneda")
                      FechaTransaccion = FrmReportes.DtaConsulta.Recordset("FechaTransaccion")
                      NumFecha = FrmReportes.DtaConsulta.Recordset("FechaTransaccion")
                      Fechas = FrmReportes.DtaConsulta.Recordset("FechaTransaccion")
                      FrmReportes.DtaTasas2.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas WHERE (FechaTasas = '" & Format(Fechas, "yyyymmdd") & "')"
                      FrmReportes.DtaTasas2.Refresh
                      If Not FrmReportes.DtaTasas2.Recordset.EOF Then
                        TasaCambio = FrmReportes.DtaTasas2.Recordset("MontoCordobas")
                      Else
                        TasaCambio = 0
                      End If
                     If TasaCambio = 0 Then
                      FrmReportes.osProgress2.Visible = False
                      FrmReportes.osProgress1.Visible = False
                      cadena = "La tasa de Cambio con Fecha: " & Fechas & vbLf
                      cadena = cadena & "no puede ser igual a Cero, el Sistema Contable" & vbLf
                      cadena = cadena & "no contiuara el proceso......"
                      MsgBox cadena, vbCritical, "Sistema Contable"
                      Exit Sub
                     End If
                      If Not IsNull(FrmReportes.DtaConsulta.Recordset("DescripcionCuentas")) Then
                       Descripcion = FrmReportes.DtaConsulta.Recordset("CodCuentas") + "." + FrmReportes.DtaConsulta.Recordset("DescripcionCuentas")
                      Else
                       Descripcion = FrmReportes.DtaConsulta.Recordset("CodCuentas") + "." + "NO TIENE DESCRIPCION????"
                      End If
                      

'                      FrmReportes.AdoConsultas.RecordSource = "SELECT  NTransaccion, FechaTransaccion, VoucherNo, ChequeNo, DescripcionMovimiento, CodCuentas, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, Debito - Credito AS Balance, TCambio, NumeroMovimiento, Beneficiario  From Transacciones  " & _
'                                                              "WHERE (FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') AND (CodCuentas = '" & CodigoCuenta & "') ORDER BY CodCuentas, FechaTransaccion, NTransaccion"

                      If FrmReportes.CmbMoneda.Text = "Córdobas" Then
                        If FrmReportes.ChkQuitarMovimiento.Value = 1 Then
'                         FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.NTransaccion, Transacciones.FechaTransaccion, Transacciones.VoucherNo, Transacciones.ChequeNo, Transacciones.DescripcionMovimiento, Transacciones.CodCuentas, Transacciones.TCambio * Transacciones.Debito AS MDebito, Transacciones.TCambio * Transacciones.Credito AS MCredito, Transacciones.Debito - Transacciones.Credito AS Balance, Transacciones.TCambio, Transacciones.NumeroMovimiento, Transacciones.Beneficiario, IndiceTransaccion.Fuente FROM  Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento  " & _
'                                                                 "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) AND (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (IndiceTransaccion.Fuente <> 'Cierre') ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"
                          FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.CodCuentas, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END, 2)) AS MDebito, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END, 2)) AS MCredito, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Tasas.MontoCordobas END, 2)) AS DebitoD, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Tasas.MontoCordobas END, 2)) " & _
                                                                 "AS CreditoD FROM Transacciones INNER JOIN  IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN Tasas ON Transacciones.FechaTransaccion = Tasas.FechaTasas WHERE  (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102))  AND (IndiceTransaccion.Fuente <> 'Cierre') GROUP BY Transacciones.CodCuentas, IndiceTransaccion.Ajuste HAVING (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (IndiceTransaccion.Ajuste <> '" & Ajuste & "') ORDER BY Transacciones.CodCuentas "
                        Else
'                        FrmReportes.AdoConsultas.RecordSource = "SELECT  NTransaccion, FechaTransaccion, VoucherNo, ChequeNo, DescripcionMovimiento, CodCuentas, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, Debito - Credito AS Balance, TCambio, NumeroMovimiento, Beneficiario From Transacciones  " & _
'                                                              "WHERE (FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') AND (CodCuentas = '" & CodigoCuenta & "') ORDER BY CodCuentas, FechaTransaccion, NTransaccion "
                         FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.CodCuentas, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END, 2)) AS MDebito, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END, 2)) AS MCredito, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Tasas.MontoCordobas END, 2)) AS DebitoD, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Tasas.MontoCordobas END, 2)) " & _
                                                                 "AS CreditoD FROM Transacciones INNER JOIN  IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN Tasas ON Transacciones.FechaTransaccion = Tasas.FechaTasas WHERE  (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas, IndiceTransaccion.Ajuste HAVING (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (IndiceTransaccion.Ajuste <> '" & Ajuste & "') ORDER BY Transacciones.CodCuentas "
                        End If
                      Else
'                          FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.CodCuentas, SUM(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas)) AS MDebito, SUM(Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas)) AS MCredito, SUM(Transacciones.Debito - Transacciones.Credito) AS Balance FROM Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
'                                                                  "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING (Transacciones.CodCuentas = '" & CodigoCuenta & "') ORDER BY Transacciones.CodCuentas"
                         FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.CodCuentas, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END, 2)) AS MDebitoC, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END, 2)) AS MCreditoC, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Tasas.MontoCordobas END, 2)) AS MDebito, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Tasas.MontoCordobas END, 2)) AS MCredito  FROM  Transacciones INNER JOIN  IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN " & _
                                                                 "Tasas ON Transacciones.FechaTransaccion = Tasas.FechaTasas WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas, IndiceTransaccion.Ajuste HAVING (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (IndiceTransaccion.Ajuste <> '" & Ajuste & "') ORDER BY Transacciones.CodCuentas"
                    End If
                      
                      
                      
                      
                      FrmReportes.AdoConsultas.Refresh
                      TotalDebito = 0
                      TotalCredito = 0
                      Do While Not FrmReportes.AdoConsultas.Recordset.EOF
                        If Not IsNull(FrmReportes.AdoConsultas.Recordset("MDebito")) Then
                            Debito = FrmReportes.AdoConsultas.Recordset("MDebito")
'                            Debito = TRUNC(Debito, 4)
'                            Debito = Format(Debito, "##,##0.00")
                            TotalDebito = Debito + TotalDebito
                        End If
                        If Not IsNull(FrmReportes.AdoConsultas.Recordset("MCredito")) Then
                            Credito = FrmReportes.AdoConsultas.Recordset("MCredito")
'                            Credito = TRUNC(Credito, 4)
'                            Credito = Format(Credito, "##,##0.00")
                            TotalCredito = Credito + TotalCredito

                        End If
                      
                        FrmReportes.AdoConsultas.Recordset.MoveNext
                      Loop
                      
                      Debito = TotalDebito
                      Credito = TotalCredito
                      
'                      If TipoCuenta = "Inventario" Then
'                       cod = 1
'                      End If
'
                                                              
                      If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                
                            If FrmReportes.ChkQuitarMovimiento.Value = 1 Then
'                                 FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.NTransaccion, Transacciones.FechaTransaccion, Transacciones.VoucherNo, Transacciones.ChequeNo, Transacciones.DescripcionMovimiento, Transacciones.CodCuentas, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Transacciones.TCambio END AS MDebito, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Transacciones.TCambio END AS MCredito, Transacciones.Debito - Transacciones.Credito AS Balance, Transacciones.TCambio, Transacciones.NumeroMovimiento, Transacciones.Beneficiario, IndiceTransaccion.Nperiodo, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Transacciones.TCambio END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Transacciones.TCambio END AS CreditoD " & _
'                                                                         "FROM Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento  " & _
'                                                                         "WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') AND (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (IndiceTransaccion.Fuente <> 'Cierre') ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"
                                  FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.CodCuentas, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END, 2))AS MDebito,SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END, 2)) AS MCredito, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Tasas.MontoCordobas END, 2)) AS DebitoD, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Tasas.MontoCordobas END, 2)) AS CreditoD FROM Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN  " & _
                                                                          "Tasas ON Transacciones.FechaTransaccion = Tasas.FechaTasas  " & _
                                                                          "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) AND (IndiceTransaccion.Fuente <> 'Cierre') AND (IndiceTransaccion.Ajuste <> '" & Ajuste & "') GROUP BY Transacciones.CodCuentas HAVING (Transacciones.CodCuentas = '" & CodigoCuenta & "') ORDER BY Transacciones.CodCuentas"
                            Else
                                
'                                 FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.CodCuentas, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END AS MDebito, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END AS MCredito, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Tasas.MontoCordobas END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Tasas.MontoCordobas END AS CreditoD  " & _
'                                                                         "FROM Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN  Tasas ON Transacciones.FechaTransaccion = Tasas.FechaTasas  " & _
'                                                                         "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) AND  (Transacciones.CodCuentas = '" & CodigoCuenta & "') ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion"
                                
                                  FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.CodCuentas, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END, 2))AS MDebito,SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END, 2)) AS MCredito, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Tasas.MontoCordobas END, 2)) AS DebitoD, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Tasas.MontoCordobas END, 2)) AS CreditoD FROM Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN  " & _
                                                                          "Tasas ON Transacciones.FechaTransaccion = Tasas.FechaTasas  " & _
                                                                          "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) AND (IndiceTransaccion.Ajuste <> '" & Ajuste & "') GROUP BY Transacciones.CodCuentas HAVING (Transacciones.CodCuentas = '" & CodigoCuenta & "') ORDER BY Transacciones.CodCuentas"
                            End If
                                                                        
                      FrmReportes.AdoConsultas.Refresh
                      TotalDebito = 0
                      TotalCredito = 0
                      Do While Not FrmReportes.AdoConsultas.Recordset.EOF
                        If Not IsNull(FrmReportes.AdoConsultas.Recordset("MDebito")) Then
                          If FrmReportes.CmbMoneda.Text = "Dólares" Then
                            Debito = FrmReportes.AdoConsultas.Recordset("DebitoD")
                          Else
                            Debito = FrmReportes.AdoConsultas.Recordset("MDebito")
                          End If
                            
                          
'                            Debito = TRUNC(Debito, 3)
'                            Debito = Format(Debito, "##,##0.00")
                            TotalDebito = Debito + TotalDebito
                        End If
                        
                        If Not IsNull(FrmReportes.AdoConsultas.Recordset("MCredito")) Then
                          If FrmReportes.CmbMoneda.Text = "Dólares" Then
                            Credito = FrmReportes.AdoConsultas.Recordset("CreditoD")
                          Else
                            Credito = FrmReportes.AdoConsultas.Recordset("MCredito")
                          End If
'                            Credito = TRUNC(Credito, 3)
'                            Credito = Format(Credito, "##,##0.00")
                            TotalCredito = Credito + TotalCredito
                        End If
                      
                        FrmReportes.AdoConsultas.Recordset.MoveNext
                      Loop
                      
                      Debito = TotalDebito
                      Credito = TotalCredito
                        
                      
                        
                        'borrar balanza,  si no funciona, totaldebito y total credito no se usan para balanza, hasta que yo
'                        TotalDebito = TotalDebito + Debito
'                        TotalCredito = TotalCredito + Credito
                        
                        '//////////////////Verifico el tipo de moneda de la cuenta////////////////////
                        If Not TipoMoneda = FrmReportes.CmbMoneda.Text Then
                           Select Case TipoMoneda
                              Case "Córdobas"
                                    TotalCuenta = Format((Debito - Credito) + TotalCuenta, "####0.00")
                          
                              Case "Dólares"
                                    TotalCuenta = Format((DebitoD - CreditoD) + TotalCuenta, "####0.00")
                           
                          End Select
                        Else
                               TotalCuenta = Format((Debito - Credito) + TotalCuenta, "####0.00")
                        End If
                        
                          Total1 = Debito - Credito + Total1
                
'                        Debito = 0
'                        Credito = 0
            Else
                      
                            If FrmReportes.ChkQuitarMovimiento.Value = 1 Then
'                                 FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.NTransaccion, Transacciones.FechaTransaccion, Transacciones.VoucherNo, Transacciones.ChequeNo, Transacciones.DescripcionMovimiento, Transacciones.CodCuentas, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Transacciones.TCambio END AS MDebito, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Transacciones.TCambio END AS MCredito, Transacciones.Debito - Transacciones.Credito AS Balance, Transacciones.TCambio, Transacciones.NumeroMovimiento, Transacciones.Beneficiario, IndiceTransaccion.Nperiodo, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Transacciones.TCambio END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Transacciones.TCambio END AS CreditoD " & _
'                                                                         "FROM Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento  " & _
'                                                                         "WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') AND (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (IndiceTransaccion.Fuente <> 'Cierre') ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"
                                  FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.CodCuentas, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END, 2))AS MDebito,SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END, 2)) AS MCredito, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Tasas.MontoCordobas END, 2)) AS DebitoD, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Tasas.MontoCordobas END, 2)) AS CreditoD FROM Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN  " & _
                                                                          "Tasas ON Transacciones.FechaTransaccion = Tasas.FechaTasas  " & _
                                                                          "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) AND (IndiceTransaccion.Fuente <> 'Cierre') AND (IndiceTransaccion.Ajuste <> '" & Ajuste & "') GROUP BY Transacciones.CodCuentas HAVING (Transacciones.CodCuentas = '" & CodigoCuenta & "') ORDER BY Transacciones.CodCuentas"
                            Else
'                                  FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.CodCuentas, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END, 2))AS MDebito,SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END, 2)) AS MCredito, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Tasas.MontoCordobas END, 2)) AS DebitoD, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Tasas.MontoCordobas END, 2)) AS CreditoD FROM Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN  " & _
'                                                                          "Tasas ON Transacciones.FechaTransaccion = Tasas.FechaTasas  " & _
'                                                                          "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING (Transacciones.CodCuentas = '" & CodigoCuenta & "') ORDER BY Transacciones.CodCuentas"
                           
                           
                                  FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.CodCuentas, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END, 2))AS MDebito,SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END, 2)) AS MCredito, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Tasas.MontoCordobas END, 2)) AS DebitoD, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Tasas.MontoCordobas END, 2)) AS CreditoD FROM Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN  " & _
                                                                          "Tasas ON Transacciones.FechaTransaccion = Tasas.FechaTasas  " & _
                                                                          "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) AND (IndiceTransaccion.Ajuste <> '" & Ajuste & "') GROUP BY Transacciones.CodCuentas HAVING (Transacciones.CodCuentas = '" & CodigoCuenta & "') ORDER BY Transacciones.CodCuentas"
                           End If
                                                                        
                      FrmReportes.AdoConsultas.Refresh
                      TotalDebito = 0
                      TotalCredito = 0
                      Do While Not FrmReportes.AdoConsultas.Recordset.EOF
                        If Not IsNull(FrmReportes.AdoConsultas.Recordset("MDebito")) Then
                          If FrmReportes.CmbMoneda.Text = "Dólares" Then
                            Debito = FrmReportes.AdoConsultas.Recordset("DebitoD")
                          Else
                            Debito = FrmReportes.AdoConsultas.Recordset("MDebito")
                          End If
'                            Debito = TRUNC(Debito, 3)
'                            Debito = Format(Debito, "##,##0.00")
                            TotalDebito = Debito + TotalDebito
                        End If
                        If Not IsNull(FrmReportes.AdoConsultas.Recordset("MCredito")) Then
                          If FrmReportes.CmbMoneda.Text = "Dólares" Then
                            Credito = FrmReportes.AdoConsultas.Recordset("CreditoD")
                          Else
                            Credito = FrmReportes.AdoConsultas.Recordset("MCredito")
                          End If
'                            Credito = TRUNC(Credito, 3)
'                            Credito = Format(Credito, "##,##0.00")
                            TotalCredito = Credito + TotalCredito
                        End If
                      
                        FrmReportes.AdoConsultas.Recordset.MoveNext
                      Loop
                      
                      Debito = TotalDebito
                      Credito = TotalCredito
                      
'                         If Not IsNull(FrmReportes.AdoConsultas.Recordset("MDebito")) Then
'                            Debito = FrmReportes.AdoConsultas.Recordset("MDebito")
'                            Debito = TRUNC(Debito, 3)
'                         End If
'                         If Not IsNull(FrmReportes.AdoConsultas.Recordset("MCredito")) Then
'                            Credito = FrmReportes.AdoConsultas.Recordset("MCredito")
'                            Credito = TRUNC(Credito, 3)
'                            Credito = Format(Credito, "##,##0.00")
'                         End If
                         
                        '//////////////////Verifico el tipo de moneda de la cuenta////////////////////
                        If Not TipoMoneda = FrmReportes.CmbMoneda.Text Then
                           Select Case TipoMoneda
                              Case "Córdobas"
                                    TotalCuenta = Format((Credito - Debito) + TotalCuenta, "####0.00")
                          
                              Case "Dólares"
                                    TotalCuenta = Format((Credito - Debito) + TotalCuenta, "####0.00")
                           
                          End Select
                        Else
                               TotalCuenta = Format((Credito - Debito) + TotalCuenta, "####0.00")
                        End If
                         
                         Total1 = Credito - Debito + Total1
'                         Debito = 0
'                         Credito = 0
                      End If
                    
                    
                
                   
                   FrmReportes.DtaConsulta.Recordset.MoveNext
                
                   Loop

         '/////////////////////////////////////////////////////////////////////////////////////////////////////////
         '////////////////////////////////GRABO LAS CUENTAS DEL PERIODO/////////////////////////////////////
         '///////////////////////////////////////////////////////////////////////////////////////////////////////


                    If QUIEN = "Balanza" Then
                '//////////////////////////Busco la Cuenta en Reportes/////////////////////////////////
                       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.KeyGrupoCuenta,Reportes.KeyGrupoSuperior,Reportes.KeyGrupo,Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3,CodCuentas From Reportes Where (((Reportes.KeyGrupo) = '" & CodigoCuenta & "'))"
                       FrmReportes.DtaConsulta2.Refresh
                '       InputBox "", "", FrmReportes.DtaConsulta2.RecordSource
                       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
                
                '          'FrmReportes.DtaConsulta2.Recordset.Edit
                          If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                           If TotalCuenta < 0 Then
                             FrmReportes.DtaConsulta2.Recordset("Haber2") = Credito
                             FrmReportes.DtaConsulta2.Recordset("Debe2") = Debito
'                             Abs(TotalCuenta)
                           Else
                             FrmReportes.DtaConsulta2.Recordset("Debe2") = Debito
                             FrmReportes.DtaConsulta2.Recordset("Haber2") = Credito
'                             TotalCuenta
                           End If
                          Else
                           If TotalCuenta < 0 Then
                               FrmReportes.DtaConsulta2.Recordset("Debe2") = Debito
                               FrmReportes.DtaConsulta2.Recordset("Haber2") = Credito
'                               Abs(TotalCuenta)
                           Else
                               FrmReportes.DtaConsulta2.Recordset("Haber2") = Credito
                               FrmReportes.DtaConsulta2.Recordset("Debe2") = Debito
'                               TotalCuenta
                           End If
                          End If
'                          FrmReportes.DtaConsulta2.Recordset("CodCuentas") = CodigoCuenta
                          FrmReportes.DtaConsulta2.Recordset.Update
                       End If
                       
                       
                       
                '//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
                              Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta"))
                              If Longitud > 1 Then
                                 If Longitud = 5 Then
                                     Nivel = 2
                                 Else
                                     Nivel = (Longitud - 5) / 2
                                     Nivel = Nivel + 2
                                 End If
                              Else
                                    Nivel = 1
                              End If
                       
                '///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
                       Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta")
                       

                       'Nivel = Nivel - 1
                       For i = Nivel To 1 Step -1
                '/////////Busco el Grupo para Sumar los Totaldes
                           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, CodCuentas From Reportes Where (((Reportes.KeyGrupo) = '" & Respuesta & "')) ORDER BY Orden"
                           FrmReportes.DtaConsulta.Refresh
                           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                              FrmReportes.DtaConsulta.Recordset.MoveLast
                '              'FrmReportes.'DtaConsulta.Recordset.Edit
                              If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                                If (TotalCuenta) < 0 Then
                                  FrmReportes.DtaConsulta.Recordset("Haber2") = FrmReportes.DtaConsulta.Recordset("Haber2") + Credito 'Abs(TotalCuenta)
                                  FrmReportes.DtaConsulta.Recordset("Debe2") = FrmReportes.DtaConsulta.Recordset("Debe2") + Debito
                                Else
                                 FrmReportes.DtaConsulta.Recordset("Debe2") = FrmReportes.DtaConsulta.Recordset("Debe2") + Debito 'TotalCuenta
                                 FrmReportes.DtaConsulta.Recordset("Haber2") = FrmReportes.DtaConsulta.Recordset("Haber2") + Credito
                                End If
                              Else
                               If TotalCuenta < 0 Then
                                 FrmReportes.DtaConsulta.Recordset("Debe2") = FrmReportes.DtaConsulta.Recordset("Debe2") + Debito 'Abs(TotalCuenta)
                                 FrmReportes.DtaConsulta.Recordset("Haber2") = FrmReportes.DtaConsulta.Recordset("Haber2") + Credito
                               Else
                                 FrmReportes.DtaConsulta.Recordset("Debe2") = FrmReportes.DtaConsulta.Recordset("Debe2") + Debito
                                 FrmReportes.DtaConsulta.Recordset("Haber2") = FrmReportes.DtaConsulta.Recordset("Haber2") + Credito 'TotalCuenta
                               End If
                              End If
'                              FrmReportes.DtaConsulta.Recordset("CodCuentas") = CodigoCuenta
                              FrmReportes.DtaConsulta.Recordset.Update
                              
                           If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
                              Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
                           End If
                           
                       If Respuesta = "G01000301" Then
                            Cod = 1
                       End If
                
                           End If
                

                       Next
                    
                    
                    
                    ElseIf QUIEN = "BalanzaCodigo" Then
                    
                     '////////////Agrego los Saldos del Periodo Seleccionado////////////////////
                '      FrmReportes.lblProgreso.Caption = "Agregando saldos al Periodo"
                '      FrmReportes.osProgress1.Value = 0
                '      FrmReportes.osProgress1.Max
                
                
                      
                      FrmReportes.DtaReportes.Recordset.AddNew
                      FrmReportes.DtaReportes.Recordset("Descripcion") = Descripcion
                      FrmReportes.DtaReportes.Recordset("CodCuentas") = CodigoCuenta 'Don guillermo
                      If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                        FrmReportes.DtaReportes.Recordset("Debe2") = TotalDebito
                        FrmReportes.DtaReportes.Recordset("Haber2") = TotalCredito
                      Else
                        FrmReportes.DtaReportes.Recordset("Debe2") = TotalDebito
                        FrmReportes.DtaReportes.Recordset("Haber2") = TotalCredito
                      End If
                      FrmReportes.DtaReportes.Recordset!Orden = Orden
                      Orden = Orden + 1
                        
                      FrmReportes.DtaReportes.Recordset.Update
                      
                    
                    ElseIf QUIEN = "Balance" Then
                    
                '//////////////////////////Busco la Cuenta en Reportes/////////////////////////////////
                       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.KeyGrupoCuenta,Reportes.KeyGrupoSuperior,Reportes.KeyGrupo,Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3 From Reportes Where (((Reportes.KeyGrupo) = '" & CodigoCuenta & "'))ORDER BY Orden "
                       FrmReportes.DtaConsulta2.Refresh
                       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
                
                          'FrmReportes.DtaConsulta2.Recordset.Edit
                            FrmReportes.DtaConsulta2.Recordset("Debe1") = TotalCuenta
                          FrmReportes.DtaConsulta2.Recordset.Update
                       End If
                       
                '//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
                              Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta"))
                              If Longitud > 1 Then
                                 If Longitud = 5 Then
                                     Nivel = 2
                                 Else
                                     Nivel = (Longitud - 5) / 2
                                     Nivel = Nivel + 2
                                 End If
                              Else
                                    Nivel = 1
                              End If
                       
                '///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
                       Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta")
                       'Nivel = Nivel - 1
                       For i = Nivel To 1 Step -1
                '/////////Busco el Grupo para Sumar los Totaldes
                           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = '" & Respuesta & "')) ORDER BY Orden"
                '          InputBox "", "", FrmReportes.DtaConsulta.RecordSource
                           FrmReportes.DtaConsulta.Refresh
                           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                              FrmReportes.DtaConsulta.Recordset.MoveLast
                '              'FrmReportes.'DtaConsulta.Recordset.Edit
                                 FrmReportes.DtaConsulta.Recordset("Haber1") = FrmReportes.DtaConsulta.Recordset("Haber1") + TotalCuenta
                              FrmReportes.DtaConsulta.Recordset.Update
                
                           End If
                
                           If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
                              Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
                           End If
                       Next
                       
                '//UTILIDAD
               ElseIf QUIEN = "UtilidadResultado" Then
                       If TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Then
                         TotalGastos = TotalGastos + TotalCuenta
                    
                       ElseIf TipoCuenta = "Ingresos - Ventas" Then
                         Totalingresos = Totalingresos + TotalCuenta
                       End If
                       
                '       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta From Reportes Where (((Reportes.Descripcion) Like '*Resultado Periodo*'))"
                       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta From Reportes Where (((Reportes.Descripcion) Like '%Resultado Periodo%'))"
                       FrmReportes.DtaConsulta2.Refresh
                       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
                           'FrmReportes.DtaConsulta2.Recordset.Edit
                               FrmReportes.DtaConsulta2.Recordset("Haber1") = Totalingresos - TotalGastos
                           FrmReportes.DtaConsulta2.Recordset.Update
                       End If
                      
                    
                ElseIf QUIEN = "Utilidad" Then
                       If TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Then
                         TotalGastos = TotalGastos + TotalCuenta
                    
                       ElseIf TipoCuenta = "Ingresos - Ventas" Then
                         Totalingresos = Totalingresos + TotalCuenta
                       End If
                       
                '       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta From Reportes Where (((Reportes.Descripcion) Like '*Resultado Periodo*'))"
                       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta From Reportes Where (((Reportes.Descripcion) Like '%Resultado Periodo%'))"
                       FrmReportes.DtaConsulta2.Refresh
                       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
                           'FrmReportes.DtaConsulta2.Recordset.Edit
                               FrmReportes.DtaConsulta2.Recordset("Debe1") = Totalingresos - TotalGastos
                           FrmReportes.DtaConsulta2.Recordset.Update
                    
                '//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
                              Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupo"))
                              If Longitud > 1 Then
                                 If Longitud = 5 Then
                                     Nivel = 2
                                 Else
                                     Nivel = (Longitud - 5) / 2
                                     Nivel = Nivel + 2
                                 End If
                              Else
                                    Nivel = 1
                              End If
                       
                '///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
                       Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupo")
                       'Nivel = Nivel - 1
                '       For I = Nivel To 1 Step -1
                '/////////Busco el Grupo para Sumar los Totaldes
                           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = 'PC')) ORDER BY Orden"
                           FrmReportes.DtaConsulta.Refresh
                           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                              FrmReportes.DtaConsulta.Recordset.MoveLast
                                 FrmReportes.DtaConsulta.Recordset("Haber1") = (Totalingresos - TotalGastos)
                              FrmReportes.DtaConsulta.Recordset.Update
                
                           End If
                           
                           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = 'C')) ORDER BY Orden"
                           FrmReportes.DtaConsulta.Refresh
                           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                              FrmReportes.DtaConsulta.Recordset.MoveLast
                                 FrmReportes.DtaConsulta.Recordset("Haber1") = (Totalingresos - TotalGastos)
                              FrmReportes.DtaConsulta.Recordset.Update
                
                           End If
                
                '           If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
                '              Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
                '           End If
                '       Next
                      End If
                    ElseIf QUIEN = "Resultado" Then
                '//////////////////////////Busco la Cuenta en Reportes/////////////////////////////////
                       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.KeyGrupoCuenta,Reportes.KeyGrupoSuperior,Reportes.KeyGrupo,Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3 From Reportes Where (((Reportes.KeyGrupo) = '" & CodigoCuenta & "'))"
                       FrmReportes.DtaConsulta2.Refresh
                       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
                
                          'FrmReportes.DtaConsulta2.Recordset.Edit
                            FrmReportes.DtaConsulta2.Recordset("Debe1") = TotalCuenta
                          FrmReportes.DtaConsulta2.Recordset.Update
                '       End If
                       
                '//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
                              Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta"))
                              If Longitud > 1 Then
                                 If Longitud = 5 Then
                                     Nivel = 2
                                 Else
                                     Nivel = (Longitud - 5) / 2
                                     Nivel = Nivel + 2
                                 End If
                              Else
                                    Nivel = 1
                              End If
                       
                '///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
                       Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta")
                       Else
                       
                       MsgBox "La cuenta Tiene Saldo y no aparece en la Estructura, Cuenta: " & CodigoCuenta, vbCritical
                        
                       End If
                       'Nivel = Nivel - 1
                       For i = Nivel To 1 Step -1
                '/////////Busco el Grupo para Sumar los Totaldes
                           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = '" & Respuesta & "')) ORDER BY Orden"
                '           InputBox "", "", DtaConsulta.RecordSource
                           FrmReportes.DtaConsulta.Refresh
                           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                              FrmReportes.DtaConsulta.Recordset.MoveLast
                              'FrmReportes.'DtaConsulta.Recordset.Edit
                                 FrmReportes.DtaConsulta.Recordset("Haber1") = FrmReportes.DtaConsulta.Recordset("Haber1") + TotalCuenta
                              FrmReportes.DtaConsulta.Recordset.Update
                
                           End If
                
                           If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
                              Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
                           End If
                       Next
                    End If
                    
                
                   FrmReportes.DtaHistorial.Recordset.MoveNext
                 
                  Loop
  
                 FrmReportes.osProgress2.Visible = False
                 
  
'*************************************************************************************************************
'*************************************************************************************************************
'*************************************************************************************************************
        '/////////////////////////////////////////////////////////////////////////////////////////////////////
        '//////////////////PERIODO ANTERIOR///////////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////////////////////
'*************************************************************************************************************
'*************************************************************************************************************
'*************************************************************************************************************

        If FrmReportes.CmbReportes.Text = "LISTA CUENTAS X COBRAR" Or FrmReportes.CmbReportes.Text = "LISTA CUENTAS X PAGAR" Then
          QUIEN = "SaldoCuentas"
          
        End If
  
        If FrmReportes.ChkBalanza.Value = False Then
          
            Debito = 0
            Credito = 0
            Totalingresos = 0
            TotalGastos = 0
            'Busco que cuentas tienen saldo
            If QUIEN = "Utilidad" Then
                Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) <'" & Format(FechaIni, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda" & vbLf
                Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Ingresos - Ventas' Or (Cuentas.TipoCuenta) = 'Costos' Or (Cuentas.TipoCuenta) = 'Gastos')) ORDER BY Cuentas.CodCuentas"
            ElseIf QUIEN = "BalanzaCodigo" Then
'                SQL = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio,5)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito,5)) AS MCredito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio,5)) - SUM(ROUND(Transacciones.Credito * Transacciones.TCambio,5)) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM  Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas " & _
'                      "WHERE (Transacciones.FechaTransaccion < '" & Format(FechaIni, "yyyymmdd") & "') GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda Having (SUM(Round(Transacciones.Debito * Transacciones.TCambio,5)) - SUM(Round(Transacciones.Credito * Transacciones.TCambio,5)) <> 0) ORDER BY Cuentas.CodCuentas"
                 Sql = "SELECT Transacciones.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 2)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito, 2)) AS MCredito, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas  " & _
                       "WHERE (Transacciones.FechaTransaccion < '" & Format(FechaIni, "yyyymmdd") & "') GROUP BY Transacciones.CodCuentas HAVING  (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "')"
            ElseIf QUIEN = "SaldoCuentas" Then
                 QUIEN = "BalanzaCodigo"
                 Sql = "SELECT Transacciones.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 2)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito, 2)) AS MCredito, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas  " & _
                       "WHERE (Transacciones.FechaTransaccion < '" & Format(FechaIni, "yyyymmdd") & "') GROUP BY Transacciones.CodCuentas HAVING  (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') ORDER BY Transacciones.CodCuentas"  'AND (MAX(Cuentas.TipoCuenta) = 'Cuentas x Cobrar')
            
            ElseIf QUIEN = "Balanza" Then
'                SQl = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) <'" & Format(FechaIni, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda" & vbLf
'                SQl = SQl & "ORDER BY Cuentas.CodCuentas"
        
                Sql = "SELECT Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.Credito * Transacciones.TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, MAX(Cuentas.KeyGrupo) AS KeyGrupo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (Transacciones.FechaTransaccion < '" & Format(FechaIni, "yyyymmdd") & "') GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda HAVING (MAX(Cuentas.KeyGrupo) BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') ORDER BY Cuentas.CodCuentas"
            
            ElseIf QUIEN = "Resultado" Then
              Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) <'" & Format(FechaIni, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda" & vbLf
              Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Ingresos - Ventas' Or (Cuentas.TipoCuenta) = 'Costos' Or (Cuentas.TipoCuenta) = 'Gastos')) ORDER BY Cuentas.CodCuentas"
            ElseIf QUIEN = "UtilidadResultado" Then
             Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) <'" & Format(FechaIni, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda " & vbLf
             Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Ingresos - Ventas' Or (Cuentas.TipoCuenta) = 'Costos' Or (Cuentas.TipoCuenta) = 'Gastos')) ORDER BY Cuentas.CodCuentas"
             FrmReportes.DtaHistorial.RecordSource = Sql
             FrmReportes.DtaHistorial.Refresh
            
            
            
            Else
              
                
                Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) <'" & Format(FechaIni, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda" & vbLf
                Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Otros Activos' Or (Cuentas.TipoCuenta) = 'Caja' Or (Cuentas.TipoCuenta) = 'Bancos' Or (Cuentas.TipoCuenta) = 'Cuentas x Cobrar' Or (Cuentas.TipoCuenta) = 'Inventario' Or (Cuentas.TipoCuenta) = 'Papeleria - Utiles' Or (Cuentas.TipoCuenta) = 'Activo Fijo' Or (Cuentas.TipoCuenta) = 'Otros Pasivos' Or (Cuentas.TipoCuenta) = 'Cuentas x Pagar' Or (Cuentas.TipoCuenta) = 'Pasivo' Or (Cuentas.TipoCuenta) = 'Capital')) ORDER BY Cuentas.CodCuentas"
            End If
            
                FrmReportes.DtaHistorial.RecordSource = Sql
        
                FrmReportes.DtaHistorial.Refresh
        
                FrmReportes.LblProgreso.Caption = "Consultando Registros del Periodo Anterior para " & QUIEN
                FrmReportes.osProgress1.Value = 0
            
            If Not FrmReportes.DtaHistorial.Recordset.EOF Then
                FrmReportes.osProgress1.Max = FrmReportes.DtaHistorial.Recordset.RecordCount
                FrmReportes.DtaHistorial.Refresh
            Else
                ' Exit Sub
            End If
        
            Do While Not FrmReportes.DtaHistorial.Recordset.EOF
                '////////Consulto los registros del periodo ANTERIOR.///////////
                FrmReportes.osProgress1.Value = FrmReportes.osProgress1.Value + 1
                CodigoCuenta = FrmReportes.DtaHistorial.Recordset("CodCuentas")
                

                    If CodigoCuenta = "1130-001-002" Then
                      CodigoCuenta = "1130-001-002"
                    End If
                
                FrmReportes.LblProgreso.Caption = "Consultando Registros del Periodo Anterior para la Cuenta " & CodigoCuenta
                DoEvents
            
'                FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, Transacciones.FechaTransaccion, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio - Transacciones.TCambio * Transacciones.Credito) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Tasas.MontoCordobas) AS MontoCordobas, MAX(Tasas.MontoLibras) AS MontoLibras, MAX(Transacciones.NTransaccion) AS NTransaccion FROM  Tasas INNER JOIN  Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Tasas.FechaTasas = Transacciones.FechaTasas GROUP BY Cuentas.CodCuentas, Transacciones.FechaTransaccion  " & _
'                                                       "HAVING (Cuentas.CodCuentas = '" & CodigoCuenta & "') AND (Transacciones.FechaTransaccion < CONVERT(DATETIME,'" & Format(FechaIni, "yyyymmdd") & "', 102)) ORDER BY Cuentas.CodCuentas"

                FrmReportes.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas,  SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END, 2)) AS MDebito, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END, 2)) AS MCredito, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Tasas.MontoCordobas END, 2)) AS DebitoD, SUM(ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Tasas.MontoCordobas END, 2)) AS CreditoD, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Tasas.MontoCordobas) AS MontoCordobas,Cuentas.DescripcionCuentas " & _
                                                       "FROM  Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN Tasas ON Transacciones.FechaTransaccion = Tasas.FechaTasas INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas WHERE (Transacciones.FechaTransaccion < CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102)) AND (IndiceTransaccion.Ajuste <> '" & Ajuste & "') GROUP BY Transacciones.CodCuentas, Cuentas.DescripcionCuentas HAVING (Transacciones.CodCuentas = '" & CodigoCuenta & "') ORDER BY Transacciones.CodCuentas"
                FrmReportes.DtaConsulta.Refresh
                
                TotalCuenta = 0
                Total1 = 0
                FrmReportes.osProgress2.Value = 0
            
                Do While Not FrmReportes.DtaConsulta.Recordset.EOF
                    FrmReportes.osProgress2.Visible = True
                    If FrmReportes.osProgress2.Value = 0 Then
                '            FrmReportes.lblProgreso.Caption = "Consultando Registros del periodo seleccionado para la cuenta " & CodigoCuenta
                        FrmReportes.osProgress2.Max = FrmReportes.DtaConsulta.Recordset.RecordCount
                        FrmReportes.osProgress2.Value = FrmReportes.osProgress2.Value + 1
        
                        DoEvents
                    Else
        '               FrmReportes.lblProgreso.Caption = "Consultando Registros del periodo seleccionado para la cuenta " & CodigoCuenta
                        FrmReportes.osProgress2.Value = FrmReportes.osProgress2.Value + 1
                        DoEvents
                    End If
        
                    TotalDebito = 0
                    TotalCredito = 0
                    TipoCuenta = FrmReportes.DtaConsulta.Recordset("TipoCuenta")
                    TipoMoneda = FrmReportes.DtaConsulta.Recordset("TipoMoneda")
'                    FechaTransaccion = FrmReportes.DtaConsulta.Recordset("FechaTransaccion")
'                    Fechas1 = FrmReportes.DtaConsulta.Recordset("FechaTransaccion")
                    TasaCambio = FrmReportes.DtaConsulta.Recordset("MontoCordobas")
                    If TasaCambio = 0 Then
                        cadena = "La tasa de Cambio de Cambio con Fecha: " & Fechas1 & vbLf
                        cadena = cadena & "no puede ser igual a Cero, el Sistema Contable" & vbLf
                        cadena = cadena & "no contiuara el proceso......"
                        MsgBox cadena, vbCritical, "Sistema Contable"
                        Exit Sub
                    End If
                    If Not IsNull(FrmReportes.DtaConsulta.Recordset("DescripcionCuentas")) Then
                     Descripcion = FrmReportes.DtaConsulta.Recordset("CodCuentas") + "." + FrmReportes.DtaConsulta.Recordset("DescripcionCuentas")
                    Else
                     Descripcion = FrmReportes.DtaConsulta.Recordset("CodCuentas") + "." + "NO TIEN DESCRIPCION???"
                    End If
                    
'                      FrmReportes.AdoConsultas.RecordSource = "SELECT CodCuentas, SUM(ROUND(Debito * TCambio, 2)) AS MDebito, SUM(ROUND(TCambio * Credito, 2)) AS MCredito From Transacciones  " & _
'                                                              "WHERE (FechaTransaccion < '" & Format(FechaIni, "yyyy-mm-dd") & "') GROUP BY CodCuentas HAVING (CodCuentas = '" & CodigoCuenta & "')"
'                      FrmReportes.AdoConsultas.Refresh
                    
                    
                    If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                     If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                         If Not IsNull(FrmReportes.DtaConsulta.Recordset("MDebito")) Then
                           
                          If FrmReportes.CmbMoneda.Text = "Dólares" Then
                            Debito = FrmReportes.DtaConsulta.Recordset("DebitoD")
                          Else
                            Debito = FrmReportes.DtaConsulta.Recordset("MDebito")
                          End If

                        End If
                        If Not IsNull(FrmReportes.DtaConsulta.Recordset("MCredito")) Then
                            
                             
                          If FrmReportes.CmbMoneda.Text = "Dólares" Then
                             Credito = FrmReportes.DtaConsulta.Recordset("CreditoD")
                          Else
                             Credito = FrmReportes.DtaConsulta.Recordset("MCredito")
                          End If

                         End If
                     End If
                        Total1 = Debito - Credito + Total1
        
                '//////////////////Verifico el tipo de moneda de la cuenta////////////////////
                If Not TipoMoneda = FrmReportes.CmbMoneda.Text Then
                   Select Case TipoMoneda
                      Case "Córdobas"
                            TotalCuenta = Format((Debito - Credito) + TotalCuenta, "####0.00")
        
                      Case "Dólares"
                            TotalCuenta = Format((Debito - Credito) + TotalCuenta, "####0.00")
        
                  End Select
                Else
                       TotalCuenta = (Debito - Credito) + TotalCuenta
                End If
        
                Debito = 0
                Credito = 0
              Else
               If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                 If Not IsNull(FrmReportes.DtaConsulta.Recordset("MDebito")) Then
                    
                     If FrmReportes.CmbMoneda.Text = "Dólares" Then
                        Debito = FrmReportes.DtaConsulta.Recordset("DebitoD")
                      Else
                        Debito = FrmReportes.DtaConsulta.Recordset("MDebito")
                      End If
                 End If
                 If Not IsNull(FrmReportes.DtaConsulta.Recordset("MCredito")) Then
                          If FrmReportes.CmbMoneda.Text = "Dólares" Then
                             Credito = FrmReportes.DtaConsulta.Recordset("CreditoD")
                          Else
                             Credito = FrmReportes.DtaConsulta.Recordset("MCredito")
                          End If
                 End If
               End If
        
                '//////////////////Verifico el tipo de moneda de la cuenta////////////////////
                If Not TipoMoneda = FrmReportes.CmbMoneda.Text Then
                   Select Case TipoMoneda
                      Case "Córdobas"
                            TotalCuenta = Format((Credito - Debito) + TotalCuenta, "####0.00")
        
                      Case "Dólares"
                            TotalCuenta = Format((Credito - Debito) + TotalCuenta, "####0.00")
        
                  End Select
                Else
                       TotalCuenta = (Credito - Debito) + TotalCuenta
                End If
        
                 Total1 = Credito - Debito + Total1
                 Debito = 0
                 Credito = 0
              End If
        
            FrmReportes.DtaConsulta.Recordset.MoveNext
        
           Loop
           
        'End If
        
        
           
            '//////////////////////////////////////////////////////////////////////
            '////////////////////GRABO LOS REGISTROS DEL PERIODO ANTERIOR//////////
            '//////////////////EN LA TABLA REPORTES////////////////////////////////
            If QUIEN = "Balanza" Then
            '//////////////////////////Busco la Cuenta en Reportes/////////////////////////////////
               FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.KeyGrupoCuenta,Reportes.KeyGrupoSuperior,Reportes.KeyGrupo,Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3,reportes.orden From Reportes Where (((Reportes.KeyGrupo) = '" & CodigoCuenta & "'))"
               FrmReportes.DtaConsulta2.Refresh
               If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
        
                  'FrmReportes.DtaConsulta2.Recordset.Edit
                  If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                    FrmReportes.DtaConsulta2.Recordset("Debe1") = TotalCuenta
'                   If TotalCuenta < 0 Then
'                     FrmReportes.DtaConsulta2.Recordset("Haber1") = Abs(TotalCuenta)
'                    Else
'                     FrmReportes.DtaConsulta2.Recordset("Debe1") = TotalCuenta
'                    End If
                  Else
                     FrmReportes.DtaConsulta2.Recordset("Haber1") = TotalCuenta
                  
'                    If TotalCuenta < 0 Then
'                      FrmReportes.DtaConsulta2.Recordset("Debe1") = Abs(TotalCuenta)
'                    Else
'                     FrmReportes.DtaConsulta2.Recordset("Haber1") = TotalCuenta
'                    End If
                  End If
        '          FrmReportes.DtaConsulta2.Recordset!Orden = Orden
                  Orden = Orden + 1
                  FrmReportes.DtaConsulta2.Recordset.Update
               End If
               
                '//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
                      Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta"))
                      If Longitud > 1 Then
                         If Longitud = 5 Then
                             Nivel = 2
                         Else
                             Nivel = (Longitud - 5) / 2
                             Nivel = Nivel + 2
                         End If
                      Else
                            Nivel = 1
                      End If
               
                '///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
               Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta")
               'Nivel = Nivel - 1
               For i = Nivel To 1 Step -1
                '/////////Busco el Grupo para Sumar los Totaldes
                   FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior,reportes.orden From Reportes Where (((Reportes.KeyGrupo) = '" & Respuesta & "')) ORDER BY Orden"
                   FrmReportes.DtaConsulta.Refresh
                   If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                      FrmReportes.DtaConsulta.Recordset.MoveLast
                      'FrmReportes.'DtaConsulta.Recordset.Edit
                      If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                         FrmReportes.DtaConsulta.Recordset("Debe1") = FrmReportes.DtaConsulta.Recordset("Debe1") + TotalCuenta

'                        If TotalCuenta < 0 Then
'                          FrmReportes.DtaConsulta.Recordset("Haber1") = FrmReportes.DtaConsulta.Recordset("Haber1") + Abs(TotalCuenta)
'                        Else
'                         FrmReportes.DtaConsulta.Recordset("Debe1") = FrmReportes.DtaConsulta.Recordset("Debe1") + TotalCuenta
'                        End If
                      Else
                         FrmReportes.DtaConsulta.Recordset("Haber1") = FrmReportes.DtaConsulta.Recordset("Haber1") + TotalCuenta
'                        If TotalCuenta < 0 Then
'                         FrmReportes.DtaConsulta.Recordset("Debe1") = FrmReportes.DtaConsulta.Recordset("Debe1") + Abs(TotalCuenta)
'                        Else
'                         FrmReportes.DtaConsulta.Recordset("Haber1") = FrmReportes.DtaConsulta.Recordset("Haber1") + TotalCuenta
'                        End If
                      End If
                      If 1 = 0 Then FrmReportes.DtaConsulta.Recordset!Orden = Orden + 1
                      Orden = Orden + 1
                      
                      FrmReportes.DtaConsulta.Recordset.Update
                      
                   If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
                      Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
                   End If
        
                   End If
        

               Next
            
            
            
            ElseIf QUIEN = "BalanzaCodigo" Then
            
              '////////////Agrego los Saldos del PeriodoAnterior////////////////////
              
              Descripcion = Replace(Descripcion, "'", "")
              
               FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3,reportes.orden,Reportes.CodCuentas From Reportes Where (((Reportes.Descripcion) = '" & Descripcion & "'))"
               FrmReportes.DtaConsulta2.Refresh
               If FrmReportes.DtaConsulta2.Recordset.EOF Then
                 FrmReportes.DtaConsulta2.Recordset.AddNew
                    FrmReportes.DtaConsulta2.Recordset("Descripcion") = Descripcion
                    FrmReportes.DtaConsulta2.Recordset("CodCuentas") = CodigoCuenta
                     FrmReportes.DtaConsulta2.Recordset!Orden = Orden
                     Orden = Orden + 1
                   Else
                     'FrmReportes.DtaConsulta2.Recordset.Edit
                   End If
                  
                   If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                     If TotalCuenta < 0 Then
                       FrmReportes.DtaConsulta2.Recordset("Haber1") = Abs(TotalCuenta)
                     Else
                      FrmReportes.DtaConsulta2.Recordset("Debe1") = TotalCuenta
                     End If
                   Else
                     If TotalCuenta < 0 Then
                      FrmReportes.DtaConsulta2.Recordset("Debe1") = Abs(TotalCuenta)
                     Else
                      FrmReportes.DtaConsulta2.Recordset("Haber1") = TotalCuenta
                     End If
                   End If
            '       FrmReportes.DtaConsulta2.Recordset!Orden = Orden
            '       Orden = Orden + 1
                   FrmReportes.DtaConsulta2.Recordset.Update
                   
               
               '////////////////////AGREGO LOS SALDOS ANTERIORES DEL BALANCE///////////
               ElseIf QUIEN = "Balance" Then
            '//////////////////////////Busco la Cuenta en Reportes/////////////////////////////////
                   FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.KeyGrupoCuenta,Reportes.KeyGrupoSuperior,Reportes.KeyGrupo,Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3 From Reportes Where (((Reportes.KeyGrupo) = '" & CodigoCuenta & "'))ORDER BY Orden "
                   FrmReportes.DtaConsulta2.Refresh
                   If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
            
                      'FrmReportes.DtaConsulta2.Recordset.Edit
                        FrmReportes.DtaConsulta2.Recordset("Debe1") = TotalCuenta + FrmReportes.DtaConsulta2.Recordset("Debe1")
                      FrmReportes.DtaConsulta2.Recordset.Update
                   End If
            
            '//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
                          Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta"))
                          If Longitud > 1 Then
                             If Longitud = 5 Then
                                 Nivel = 2
                             Else
                                 Nivel = (Longitud - 5) / 2
                                 Nivel = Nivel + 2
                             End If
                          Else
                                Nivel = 1
                          End If
            
            '///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
                   Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta")
                   'Nivel = Nivel - 1
                   For i = Nivel To 1 Step -1
            '/////////Busco el Grupo para Sumar los Totaldes
                       FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = '" & Respuesta & "')) ORDER BY Orden"
            '          InputBox "", "", FrmReportes.DtaConsulta.RecordSource
                       FrmReportes.DtaConsulta.Refresh
                       If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                          FrmReportes.DtaConsulta.Recordset.MoveLast
            '              'FrmReportes.'DtaConsulta.Recordset.Edit
                             FrmReportes.DtaConsulta.Recordset("Haber1") = FrmReportes.DtaConsulta.Recordset("Haber1") + TotalCuenta
                          FrmReportes.DtaConsulta.Recordset.Update
            
                       End If
            
                       If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
                          Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
                       End If
                   Next
               
             
             '//////////////////AGREGO LA UTILIDAD AL BALANCE//////////////////////////
             ElseIf QUIEN = "Utilidad" Then
                   If TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Then
                     TotalGastos = TotalCuenta
                     Totalingresos = 0
                   ElseIf TipoCuenta = "Ingresos - Ventas" Then
                     Totalingresos = TotalCuenta
                     TotalGastos = 0
                   End If
                   
                   FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta From Reportes Where (((Reportes.Descripcion) Like '%Resultado Periodo%'))"
                   FrmReportes.DtaConsulta2.Refresh
                   If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
                           FrmReportes.DtaConsulta2.Recordset("Debe1") = Totalingresos - TotalGastos + FrmReportes.DtaConsulta2.Recordset("Debe1")
                       FrmReportes.DtaConsulta2.Recordset.Update
                
            '//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
                          Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupo"))
                          If Longitud > 1 Then
                             If Longitud = 5 Then
                                 Nivel = 2
                             Else
                                 Nivel = (Longitud - 5) / 2
                                 Nivel = Nivel + 2
                             End If
                          Else
                                Nivel = 1
                          End If
                   
            '///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
                   Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupo")
                   'Nivel = Nivel - 1
            '       For I = Nivel To 1 Step -1
            '/////////Busco el Grupo para Sumar los Totaldes
                       FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = 'PC')) ORDER BY Orden"
                       FrmReportes.DtaConsulta.Refresh
                       If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                          FrmReportes.DtaConsulta.Recordset.MoveLast
                             FrmReportes.DtaConsulta.Recordset("Haber1") = (Totalingresos - TotalGastos) + FrmReportes.DtaConsulta.Recordset("Haber1")
                          FrmReportes.DtaConsulta.Recordset.Update
            
                       End If
                       
                       FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = 'C')) ORDER BY Orden"
                       FrmReportes.DtaConsulta.Refresh
                       If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                          FrmReportes.DtaConsulta.Recordset.MoveLast
                             FrmReportes.DtaConsulta.Recordset("Haber1") = (Totalingresos - TotalGastos) + FrmReportes.DtaConsulta.Recordset("Haber1")
                          FrmReportes.DtaConsulta.Recordset.Update
            
                       End If
            
            '           If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
            '              Resp6.34uesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
            '           End If
            '       Next
                  End If
              
             ElseIf QUIEN = "UtilidadResultado" Then
                   If TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Then
                     TotalGastos = TotalCuenta
                     Totalingresos = 0
                   ElseIf TipoCuenta = "Ingresos - Ventas" Then
                     Totalingresos = TotalCuenta
                     TotalGastos = 0
                   End If
                   
                   FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta From Reportes Where (((Reportes.Descripcion) Like '%Resultado Periodo%'))"
                   FrmReportes.DtaConsulta2.Refresh
                   If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
                           FrmReportes.DtaConsulta2.Recordset("Haber1") = Totalingresos - TotalGastos + FrmReportes.DtaConsulta2.Recordset("Haber1")
                       FrmReportes.DtaConsulta2.Recordset.Update
                   End If
              
             
             
              ElseIf QUIEN = "Resultado" Then
                
                      '//////////////////////////Busco la Cuenta en Reportes/////////////////////////////////
                       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.KeyGrupoCuenta,Reportes.KeyGrupoSuperior,Reportes.KeyGrupo,Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3 From Reportes Where (((Reportes.KeyGrupo) = '" & CodigoCuenta & "'))"
                       FrmReportes.DtaConsulta2.Refresh
                       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
                
                          'FrmReportes.DtaConsulta2.Recordset.Edit
                            FrmReportes.DtaConsulta2.Recordset("Debe1") = TotalCuenta + FrmReportes.DtaConsulta2.Recordset("Debe1")
                          FrmReportes.DtaConsulta2.Recordset.Update
                '       End If
                       
                        '//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
                              Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta"))
                              If Longitud > 1 Then
                                 If Longitud = 5 Then
                                     Nivel = 2
                                 Else
                                     Nivel = (Longitud - 5) / 2
                                     Nivel = Nivel + 2
                                 End If
                              Else
                                    Nivel = 1
                              End If
                       
                         '///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
                         Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta")
                       Else
                       
                         MsgBox "La cuenta Tiene Saldo y no aparece en la Estructura, Cuenta: " & CodigoCuenta, vbCritical
                        
                       End If
                       'Nivel = Nivel - 1
                       For i = Nivel To 1 Step -1
                           '/////////Busco el Grupo para Sumar los Totaldes
                           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = '" & Respuesta & "')) ORDER BY Orden"
                '           InputBox "", "", DtaConsulta.RecordSource
                           FrmReportes.DtaConsulta.Refresh
                           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                              FrmReportes.DtaConsulta.Recordset.MoveLast
                              'FrmReportes.'DtaConsulta.Recordset.Edit
                                 FrmReportes.DtaConsulta.Recordset("Haber1") = FrmReportes.DtaConsulta.Recordset("Haber1") + TotalCuenta
                              FrmReportes.DtaConsulta.Recordset.Update
                
                           End If
                
                           If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
                              Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
                           End If
                         Next



            
                End If
            
               
               FrmReportes.DtaHistorial.Recordset.MoveNext
              Loop
              
            End If '////////FIN DEL IF PARA If FrmReportes.ChkBalanza.Value = True Then
            
            
              
            '/////////////////////////////////////////////////////////////////////////////////////////
            '//////////////////////SUMO LOS TOTALES DEL PASIVO + CAPITAL/////////////////////////////
            '/////////////////////////////////////////////////////////////////////////////////////////
              
            If QUIEN = "Balance" Then
                  FrmReportes.DtaConsulta2.RecordSource = "SELECT Sum(Reportes.Debe1) AS SumaDeDebe1, Sum(Reportes.Haber1) AS SumaDeHaber1 From Reportes Where (((Reportes.Descripcion) Like 'Total%') And ((Reportes.KeyGrupo) = 'B' Or (Reportes.KeyGrupo) = 'C'))"
            
               FrmReportes.DtaConsulta2.Refresh
               If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
                  FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.KeyGrupo From Reportes Where (((Reportes.KeyGrupo) = 'PC'))"
                  FrmReportes.DtaConsulta.Refresh
                  If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                   If Not IsNull(FrmReportes.DtaConsulta2.Recordset("SumaDeHaber1")) Then
                    TotalCuenta = FrmReportes.DtaConsulta2.Recordset("SumaDeHaber1")
                    'FrmReportes.'DtaConsulta.Recordset.Edit
                     FrmReportes.DtaConsulta.Recordset("Haber1") = TotalCuenta
                    FrmReportes.DtaConsulta.Recordset.Update
                   End If
                  End If
               End If
               
             ElseIf QUIEN = "Resultado" Then
                FrmReportes.DtaConsulta2.RecordSource = "SELECT Sum(Reportes.Debe1) AS SumaDeDebe1, Sum(Reportes.Haber1) AS SumaDeHaber1 From Reportes Where (((Reportes.Descripcion) Like 'Total%') And ((Reportes.KeyGrupo) = 'G' Or (Reportes.KeyGrupo) = 'O'))"
               FrmReportes.DtaConsulta2.Refresh
               If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
                  FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.KeyGrupo From Reportes Where (((Reportes.KeyGrupo) = 'CG'))"
                  FrmReportes.DtaConsulta.Refresh
                  If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                   If Not IsNull(FrmReportes.DtaConsulta2.Recordset("SumaDeHaber1")) Then
                    TotalCuenta = FrmReportes.DtaConsulta2.Recordset("SumaDeHaber1")
                    'FrmReportes.'DtaConsulta.Recordset.Edit
                     FrmReportes.DtaConsulta.Recordset("Haber1") = TotalCuenta
                    FrmReportes.DtaConsulta.Recordset.Update
                   End If
                  End If
               End If
             End If
  
     '////////////Agrego los Saldos de los acumulados////////////////////
'   FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.Debe1+Reportes.Debe2 AS TotalDebe, Reportes.Haber1+Reportes.Haber2 AS TotalHaber From Reportes"
'   FrmReportes.DtaConsulta2.Refresh
'
'   FrmReportes.LblProgreso.Caption = "Agregando Saldos Acumulados a las Cuentas"
'    FrmReportes.osProgress1.Value = 0
'    FrmReportes.osProgress1.Max = FrmReportes.DtaConsulta2.Recordset.RecordCount
'
'  Do While Not FrmReportes.DtaConsulta2.Recordset.EOF
'      FrmReportes.osProgress1.Value = FrmReportes.osProgress1.Value + 1
'      FrmReportes.DtaConsulta2.Recordset("Debe3") = FrmReportes.DtaConsulta2.Recordset("TotalDebe")
'      FrmReportes.DtaConsulta2.Recordset("Haber3") = FrmReportes.DtaConsulta2.Recordset("TotalHaber")
'
'    FrmReportes.DtaConsulta2.Recordset.Update
'   FrmReportes.DtaConsulta2.Recordset.MoveNext
'   Loop
'
'   FrmReportes.DtaConsulta2.Refresh
    
Ejecutar.Execute "Update Reportes set debe3=debe1+debe2,haber3=haber1+haber2"
 
End Sub





Public Sub SaldoReportesDpto(QUIEN As String)
Dim CodigoGrupo As String, Sql As String, Fechas As Date
Dim Nivel As Integer, Longitud As Integer, Fecha1 As String
Dim TotalMayor() As String, TotalDescripcion As String
Dim KeySuperior As String, NumeroHijos As Double, NumeroHijosTotales As Double
Dim DescripCuenta As String, DescripcionPadre As String, KeyUltimo As String, CodigoCuentaDesde As String, CodigoCuentaHasta As String
Dim DebitoD As Double, CreditoD As Double, CodDepartamento As String
Dim TotalDebitoDpto As Double, TotalCreditoDpto As Double


  '   ////////////////Elimino los registros del reporte///////////////////
  'frmreportes.DtaElimina.RecordSource = "DELETE Reportes.* From Reportes"
  'frmreportes.DtaElimina.Recordset.Updatable
  
    Dim Orden As Integer  'sirve para ordenar las cuentas
                Orden = 1
             
                NumFecha1 = FechaIni
                NumFecha2 = FechaFin
            
                'Busco que cuentas tienen saldo
                If QUIEN = "Balance" Then
                 Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda " & vbLf
                 Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Otros Activos' Or (Cuentas.TipoCuenta) = 'Caja' Or (Cuentas.TipoCuenta) = 'Bancos' Or (Cuentas.TipoCuenta) = 'Cuentas x Cobrar' Or (Cuentas.TipoCuenta) = 'Inventario' Or (Cuentas.TipoCuenta) = 'Papeleria - Utiles' Or (Cuentas.TipoCuenta) = 'Activo Fijo' Or (Cuentas.TipoCuenta) = 'Otros Pasivos' Or (Cuentas.TipoCuenta) = 'Cuentas x Pagar' Or (Cuentas.TipoCuenta) = 'Pasivo' Or (Cuentas.TipoCuenta) = 'Capital')) ORDER BY Cuentas.CodCuentas"
                 FrmReportes.DtaHistorial.RecordSource = Sql
                 FrmReportes.DtaHistorial.Refresh
                ElseIf QUIEN = "Utilidad" Then
                 Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda " & vbLf
                 Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Ingresos - Ventas' Or (Cuentas.TipoCuenta) = 'Costos' Or (Cuentas.TipoCuenta) = 'Gastos')) ORDER BY Cuentas.CodCuentas"
                 FrmReportes.DtaHistorial.RecordSource = Sql
                 FrmReportes.DtaHistorial.Refresh
                ElseIf QUIEN = "Resultado" Then

                    If FrmReportes.TxtDptoDesde.Text = "" And FrmReportes.TxtDptoHasta.Text = "" Then
                        Sql = "SELECT Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.Credito * Transacciones.TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta , Cuentas.TipoMoneda, Transacciones.VoucherNo,Cuentas.KeyGrupo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
                              "WHERE (Transacciones.FechaTransaccion BETWEEN  '" & Format(FechaIni, "yyyymmdd") & "' AND  '" & Format(FechaFin, "yyyymmdd") & "') GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Transacciones.VoucherNo,Cuentas.KeyGrupo HAVING (Cuentas.TipoCuenta = 'Ingresos - Ventas') OR (Cuentas.TipoCuenta = 'Costos') OR (Cuentas.TipoCuenta = 'Gastos') ORDER BY Cuentas.CodCuentas, Transacciones.VoucherNo"
                    Else
                        Sql = "SELECT Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.Credito * Transacciones.TCambio) AS Total, Cuentas.DescripcionCuentas,  Cuentas.TipoCuenta, Cuentas.TipoMoneda, Cuentas.KeyGrupo, ((CASE WHEN Transacciones.VoucherNo = '-' THEN '00' ELSE Transacciones.VoucherNo END)) AS VoucherNo  FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
                              "WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') AND (((CASE WHEN Transacciones.VoucherNo = '-' THEN '00' ELSE Transacciones.VoucherNo END)) BETWEEN '" & FrmReportes.TxtDptoDesde.Text & "' AND '" & FrmReportes.TxtDptoHasta.Text & "') GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Transacciones.VoucherNo, Cuentas.KeyGrupo HAVING (Cuentas.TipoCuenta = 'Ingresos - Ventas') OR (Cuentas.TipoCuenta = 'Costos') OR (Cuentas.TipoCuenta = 'Gastos') ORDER BY Cuentas.CodCuentas"
                    End If
                        FrmReportes.DtaHistorial.RecordSource = Sql
                        FrmReportes.DtaHistorial.Refresh
                 
                ElseIf QUIEN = "Balanza" Then
                
                     If FrmReportes.TxtDesde.Text = "" Then
                       FrmReportes.DtaConsulta.RecordSource = "SELECT * From Grupos ORDER BY KeyGrupo"
                       FrmReportes.DtaConsulta.Refresh
                       If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                         FrmReportes.DtaConsulta.Recordset.MoveFirst
                         CodigoCuentaDesde = FrmReportes.DtaConsulta.Recordset("KeyGrupo")
                       End If
                    Else
                        CodigoCuentaDesde = FrmReportes.TxtKeyGrupoDesde.Text
                    End If
                       
                    If FrmReportes.TxtHasta.Text = "" Then
                       FrmReportes.DtaConsulta.RecordSource = "SELECT * From Grupos ORDER BY KeyGrupo"
                       FrmReportes.DtaConsulta.Refresh
                       If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                         FrmReportes.DtaConsulta.Recordset.MoveLast
                         CodigoCuentaHasta = FrmReportes.DtaConsulta.Recordset("KeyGrupo")
                       End If
                    Else
                       CodigoCuentaHasta = FrmReportes.TxtKeyGrupoHasta.Text
                    End If
                
                
                
                    FrmReportes.DtaHistorial.RecordSource = "SELECT  Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 5)) AS MDebito, SUM(ROUND(Transacciones.Credito * Transacciones.TCambio,5)) AS MCredito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 5) - ROUND(Transacciones.Credito * Transacciones.TCambio, 5)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.KeyGrupo) As KeyGrupo FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
                                                            "WHERE  (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') GROUP BY Cuentas.CodCuentas HAVING  (MAX(Cuentas.KeyGrupo) BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') ORDER BY Cuentas.CodCuentas"
                    FrmReportes.DtaHistorial.Refresh
                 
                    If FrmReportes.CmbMoneda.Text = "Córdobas" Then
                        If FrmReportes.ChkQuitarMovimiento.Value = 1 Then

                          ConsultaTotalesMovimientos = "SELECT Transacciones.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 3)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito, 3)) AS MCredito, MAX(IndiceTransaccion.Fuente) AS Fuente, MAX(Cuentas.KeyGrupo) AS KeyGrupo FROM  Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING (MAX(IndiceTransaccion.Fuente) <> 'Cierre') AND (MAX(Cuentas.KeyGrupo) BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') ORDER BY Transacciones.CodCuentas"
                        Else
                          ConsultaTotalesMovimientos = "SELECT Transacciones.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 3)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito, 3)) AS MCredito, MAX(Cuentas.KeyGrupo) AS KeyGrupo FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING (MAX(Cuentas.KeyGrupo) BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "')"
                        End If
'                   Else
                        ConsultaTotalesMovimientos = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas,5)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas,5)) AS MCredito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas,5) - ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas,5)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN  Tasas ON Transacciones.FechaTransaccion = Tasas.FechaTasas  " & _
                                                     "WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyy-mm-dd") & "' AND '" & Format(FechaFin, "yyyy-mm-dd") & "') GROUP BY Cuentas.CodCuentas"
                          
                    End If
                 
                ElseIf QUIEN = "BalanzaCodigo" Then
                
                            If FrmReportes.DBCodigo.Text = "" Then
                                 FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.* From Cuentas ORDER BY CodCuentas"
                                 FrmReportes.DtaConsulta.Refresh
                                 If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                                   FrmReportes.DtaConsulta.Recordset.MoveFirst
                                   CodigoCuentaDesde = FrmReportes.DtaConsulta.Recordset("CodCuentas")
                                End If
                            Else
                              CodigoCuentaDesde = FrmReportes.DBCodigo.Text
                            End If
                            
                            If FrmReportes.DBCodigoHasta.Text = "" Then
                                 FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.* From Cuentas ORDER BY CodCuentas"
                                 FrmReportes.DtaConsulta.Refresh
                                 If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                                   FrmReportes.DtaConsulta.Recordset.MoveLast
                                   CodigoCuentaHasta = FrmReportes.DtaConsulta.Recordset("CodCuentas")
                                End If
                            Else
                               CodigoCuentaHasta = FrmReportes.DBCodigoHasta.Text
                            End If
                            
                              
                             FrmReportes.DtaHistorial.RecordSource = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 5)) AS MDebito, SUM(ROUND(Transacciones.Credito * Transacciones.TCambio,5)) AS MCredito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 5) - ROUND(Transacciones.Credito * Transacciones.TCambio, 5)) AS Total,MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Cuentas.TipoCuenta) AS TipoCuenta FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') GROUP BY Cuentas.CodCuentas HAVING (Cuentas.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') ORDER BY Cuentas.CodCuentas"
                             FrmReportes.DtaHistorial.Refresh
                                '///////////////////////////////////////////////////////////////////////////////
                                '////////////////guardo la consulta para actualizar en el reporte///////////////
                                '///////////////////////////////////////////////////////////////////////////////
                                
                                If FrmReportes.CmbMoneda.Text = "Córdobas" Then
                                    If FrmReportes.ChkQuitarMovimiento.Value = 1 Then
                                      ConsultaTotalesMovimientos = "SELECT Transacciones.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 3)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito, 3)) AS MCredito, MAX(IndiceTransaccion.Fuente) AS Fuente FROM  Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING (MAX(IndiceTransaccion.Fuente) <> 'Cierre') AND (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') ORDER BY Transacciones.CodCuentas"
                                    Else
                                      ConsultaTotalesMovimientos = "SELECT CodCuentas, SUM(ROUND(Debito * TCambio, 3)) AS MDebito, SUM(ROUND(TCambio * Credito, 3)) AS MCredito From Transacciones  WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY CodCuentas HAVING (CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') "
                                    End If
                                Else
                                     ConsultaTotalesMovimientos = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 5)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas, 5)) AS MCredito,SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 5) - ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas, 5)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTransaccion = Tasas.FechaTasas " & _
                                                                 "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Cuentas.CodCuentas HAVING (Cuentas.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') "
                                End If
                
                
                ElseIf QUIEN = "SaldoCuentas" Then
                            QUIEN = "BalanzaCodigo"
                            If FrmReportes.DBCodigo.Text = "" Then
                                 FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.* From Cuentas ORDER BY CodCuentas"
                                 FrmReportes.DtaConsulta.Refresh
                                 If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                                   FrmReportes.DtaConsulta.Recordset.MoveFirst
                                   CodigoCuentaDesde = FrmReportes.DtaConsulta.Recordset("CodCuentas")
                                End If
                            Else
                              CodigoCuentaDesde = FrmReportes.DBCodigo.Text
                            End If
                            
                            If FrmReportes.DBCodigoHasta.Text = "" Then
                                 FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.* From Cuentas ORDER BY CodCuentas"
                                 FrmReportes.DtaConsulta.Refresh
                                 If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                                   FrmReportes.DtaConsulta.Recordset.MoveLast
                                   CodigoCuentaHasta = FrmReportes.DtaConsulta.Recordset("CodCuentas")
                                End If
                            Else
                               CodigoCuentaHasta = FrmReportes.DBCodigoHasta.Text
                            End If
                            
                              
                             FrmReportes.DtaHistorial.RecordSource = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 5)) AS MDebito, SUM(ROUND(Transacciones.Credito * Transacciones.TCambio,5)) AS MCredito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 5) - ROUND(Transacciones.Credito * Transacciones.TCambio, 5)) AS Total,MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Cuentas.TipoCuenta) AS TipoCuenta FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') GROUP BY Cuentas.CodCuentas HAVING (Cuentas.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "')  ORDER BY Cuentas.CodCuentas"  'AND (MAX(Cuentas.TipoCuenta) = 'Cuentas x Cobrar')
                             FrmReportes.DtaHistorial.Refresh
                                '///////////////////////////////////////////////////////////////////////////////
                                '////////////////guardo la consulta para actualizar en el reporte///////////////
                                '///////////////////////////////////////////////////////////////////////////////
                                
                                If FrmReportes.CmbMoneda.Text = "Córdobas" Then
                                    If FrmReportes.ChkQuitarMovimiento.Value = 1 Then
                                      ConsultaTotalesMovimientos = "SELECT Transacciones.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 3)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito, 3)) AS MCredito, MAX(IndiceTransaccion.Fuente) AS Fuente FROM  Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING (MAX(IndiceTransaccion.Fuente) <> 'Cierre') AND (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') ORDER BY Transacciones.CodCuentas"
                                    Else
                                      ConsultaTotalesMovimientos = "SELECT CodCuentas, SUM(ROUND(Debito * TCambio, 3)) AS MDebito, SUM(ROUND(TCambio * Credito, 3)) AS MCredito From Transacciones  WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY CodCuentas HAVING (CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') "
                                    End If
                                Else
                                     ConsultaTotalesMovimientos = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 5)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas, 5)) AS MCredito,SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 5) - ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas, 5)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTransaccion = Tasas.FechaTasas " & _
                                                                 "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Cuentas.CodCuentas HAVING (Cuentas.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') "
                                End If
                
                
                ElseIf QUIEN = "UtilidadResultado" Then
                
                 If FrmReportes.TxtDptoDesde.Text = "" And FrmReportes.TxtDptoHasta.Text = "" Then
                    Sql = "SELECT Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.Credito * Transacciones.TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta , Cuentas.TipoMoneda, Transacciones.VoucherNo, MAX(Cuentas.KeyGrupo) AS KeyGrupo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
                          "WHERE (Transacciones.FechaTransaccion BETWEEN  '" & Format(FechaIni, "yyyymmdd") & "' AND  '" & Format(FechaFin, "yyyymmdd") & "') GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Transacciones.VoucherNo HAVING (Cuentas.TipoCuenta = 'Ingresos - Ventas') OR (Cuentas.TipoCuenta = 'Costos') OR (Cuentas.TipoCuenta = 'Gastos') ORDER BY Cuentas.CodCuentas, Transacciones.VoucherNo"
                 Else
                    Sql = "SELECT Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.Credito * Transacciones.TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Transacciones.VoucherNo, MAX(Cuentas.KeyGrupo) AS KeyGrupo FROM Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas " & _
                          "WHERE     (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') AND (Transacciones.VoucherNo BETWEEN '" & FrmReportes.TxtDptoDesde.Text & "' AND '" & FrmReportes.TxtDptoHasta.Text & "') GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Transacciones.VoucherNo HAVING (Cuentas.TipoCuenta = 'Ingresos - Ventas') OR (Cuentas.TipoCuenta = 'Costos') OR (Cuentas.TipoCuenta = 'Gastos') ORDER BY Cuentas.CodCuentas, Transacciones.VoucherNo"
                 End If
                 FrmReportes.DtaHistorial.RecordSource = Sql
                 FrmReportes.DtaHistorial.Refresh
              
                
                Else
                     FrmReportes.DtaHistorial.RecordSource = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio,5)) AS MDebito, SUM(ROUND(Transacciones.Credito * Transacciones.TCambio,5)) AS MCredito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio,5) - ROUND(Transacciones.Credito * Transacciones.TCambio,5)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Cuentas.TipoCuenta) AS TipoCuenta FROM  Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
                                                             "WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') GROUP BY Cuentas.CodCuentas ORDER BY Cuentas.CodCuentas"
                     FrmReportes.DtaHistorial.Refresh
                        '///////////////////////////////////////////////////////////////////////////////
                        '////////////////guardo la consulta para actualizar en el reporte///////////////
                        '///////////////////////////////////////////////////////////////////////////////
                        
                        If FrmReportes.CmbMoneda.Text = "Córdobas" Then
    '                         ConsultaTotalesMovimientos = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio,5)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito,5)) AS MCredito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio,5) - ROUND(Transacciones.TCambio * Transacciones.Credito,5)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda FROM Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
    '                                                      "WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') GROUP BY Cuentas.CodCuentas"
                            If FrmReportes.ChkQuitarMovimiento.Value = 1 Then
                              ConsultaTotalesMovimientos = "SELECT  Transacciones.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 3)) AS MDebito,  SUM(ROUND(Transacciones.TCambio * Transacciones.Credito, 3)) AS MCredito, MAX(IndiceTransaccion.Fuente) AS Fuente FROM  Transacciones INNER JOIN  IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento  " & _
                                                           "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING (MAX(IndiceTransaccion.Fuente) <> 'Cierre') ORDER BY Transacciones.CodCuentas"
                            Else
                              ConsultaTotalesMovimientos = "SELECT CodCuentas, SUM(ROUND(Debito * TCambio, 3)) AS MDebito, SUM(ROUND(TCambio * Credito, 3)) AS MCredito From Transacciones  WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY CodCuentas"
                            End If
    '                          ConsultaTotalesMovimientos = "SELECT   NTransaccion, FechaTransaccion, VoucherNo, ChequeNo, DescripcionMovimiento, CodCuentas, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, Debito - Credito AS Balance, TCambio, NumeroMovimiento, Beneficiario From Transacciones WHERE (FechaTransaccion BETWEEN '" & Format(FechaFin, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') ORDER BY CodCuentas, FechaTransaccion, NTransaccion"
                        Else
                            ConsultaTotalesMovimientos = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas,5)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas,5)) AS MCredito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas,5) - ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas,5)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN  Tasas ON Transacciones.FechaTransaccion = Tasas.FechaTasas  " & _
                                                         "WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyy-mm-dd") & "' AND '" & Format(FechaFin, "yyyy-mm-dd") & "') GROUP BY Cuentas.CodCuentas"
                              
                        End If
            
                End If

               Totalingresos = 0
               TotalGastos = 0
               FrmReportes.osProgress1.Value = 0
               FrmReportes.osProgress1.Visible = True
            
               If Not FrmReportes.DtaHistorial.Recordset.EOF Then
                FrmReportes.DtaHistorial.Recordset.MoveLast
                FrmReportes.osProgress1.Max = FrmReportes.DtaHistorial.Recordset.RecordCount
                FrmReportes.DtaHistorial.Recordset.MoveFirst
               End If
               
'*************************************************************************************************************
'*************************************************************************************************************
'*************************************************************************************************************
        '/////////////////////////////////////////////////////////////////////////////////////////////////////
        '//////////////////MOVIMIENTOS DEL PERIODO////////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////////////////////
'*************************************************************************************************************
'*************************************************************************************************************
'*************************************************************************************************************
 
         Do While Not FrmReportes.DtaHistorial.Recordset.EOF
                   NumFecha1 = FechaIni
                   NumFecha2 = FechaFin
                    '////////Consulto los registros del periodo seleccionado.///////////
                    CodigoCuenta = FrmReportes.DtaHistorial.Recordset("CodCuentas")
                    If QUIEN = "Resultado" Or QUIEN = "UtilidadResultado" Then
                     If Not IsNull(FrmReportes.DtaHistorial.Recordset("VoucherNo")) Then
                      CodDepartamento = FrmReportes.DtaHistorial.Recordset("VoucherNo")
                      Else
                      CodDepartamento = "-"
                     End If
                      CodigoGrupo = FrmReportes.DtaHistorial.Recordset("KeyGrupo")
                    End If
                    
                    If CodigoCuenta = "51110001" Then
                       CodigoCuenta = "51110001"
                    End If
                    
                    
                    TotalDebitoDpto = 0
                    TotalCreditoDpto = 0

                    
                    
                     FrmReportes.LblProgreso.Caption = "Consultando Registros del periodo seleccionado para la cuenta " & CodigoCuenta
                     FrmReportes.osProgress1.Value = FrmReportes.osProgress1.Value + 1
                     DoEvents
                     If FrmReportes.ChkQuitarMovimiento.Value = 1 Then
                       FrmReportes.DtaConsulta.RecordSource = "SELECT  Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio - Transacciones.TCambio * Transacciones.Credito) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NTransaccion) AS Transaccion, MAX(IndiceTransaccion.Fuente) AS Fuente  FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN  IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion  " & _
                                                              "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Cuentas.CodCuentas HAVING  (Cuentas.CodCuentas = '" & CodigoCuenta & "') AND (MAX(IndiceTransaccion.Fuente) <> 'Cierre')"
                     Else
'                       FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio - Transacciones.TCambio * Transacciones.Credito) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NTransaccion) As Transaccion FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
'                                                           "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Cuentas.CodCuentas HAVING (Cuentas.CodCuentas = '" & CodigoCuenta & "')"
                                            
                       FrmReportes.DtaConsulta.RecordSource = "SELECT  Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio - Transacciones.TCambio * Transacciones.Credito) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NTransaccion) AS Transaccion, Transacciones.VoucherNo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
                                                              "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyymmdd") & "', 102)) GROUP BY Cuentas.CodCuentas, Transacciones.VoucherNo HAVING  (Cuentas.CodCuentas = '" & CodigoCuenta & "') AND (Transacciones.VoucherNo = '" & CodDepartamento & "')"
                     End If


                    FrmReportes.DtaConsulta.Refresh
                     
                    TotalDebitoH = 0
                    TotalCreditoH = 0
                   If FrmReportes.ChkBalanza.Value = 1 Then
                   
                    '////////////////////////////////////////////////////////////////////////////////////////////
                    '///////CON ESTA CONSULTA BUSCO EL HISTORICO DEL PERIODO////////////////////////////////////////
                    '//////////////////////////////////////////////////////////////////////////////////////////////
                     
                     TotalDebitoH = 0
                     TotalCreditoH = 0
                     If FrmReportes.ChkBalanza.Value = 1 Then
'                         FrmReportes.AdoHistorial.RecordSource = "SELECT Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.Credito * Transacciones.TCambio) AS Total,Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, MAX(Transacciones.NTransaccion) AS Transaccion FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (Transacciones.FechaTransaccion < '" & Format(FechaIni, "yyyymmdd") & "') GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda HAVING (Cuentas.CodCuentas = '" & CodigoCuenta & "') ORDER BY Cuentas.CodCuentas"
                        FrmReportes.AdoHistorial.RecordSource = "SELECT Cuentas.CodCuentas, ROUND(Transacciones.Debito * Transacciones.TCambio,5) AS MDebito, ROUND(Transacciones.TCambio * Transacciones.Credito,5) AS MCredito, ROUND(Transacciones.Debito * Transacciones.TCambio,5) - ROUND(Transacciones.TCambio * Transacciones.Credito,5) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Transacciones.NTransaccion AS Transaccion FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
                                                                "WHERE (Transacciones.FechaTransaccion < '" & Format(FechaIni, "yyyy-mm-dd") & "') AND (Cuentas.CodCuentas = '" & CodigoCuenta & "') ORDER BY Cuentas.CodCuentas"
                        FrmReportes.AdoHistorial.Refresh
                        If Not FrmReportes.AdoHistorial.Recordset.EOF Then
                            If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                                If Not IsNull(FrmReportes.AdoHistorial.Recordset("MDebito")) Then
                                    TotalDebitoH = FrmReportes.AdoHistorial.Recordset("MDebito")
                                End If
                                If Not IsNull(FrmReportes.AdoHistorial.Recordset("MCredito")) Then
                                     TotalCreditoH = FrmReportes.AdoHistorial.Recordset("MCredito")
                                End If
                                            
                                TotalHistorico = TotalDebitoH - TotalCreditoH
                             Else
                               '/////////EN CASO QUE NO SEA CUENTA DE ACITVO/////
                                If Not IsNull(FrmReportes.AdoHistorial.Recordset("MDebito")) Then
                                    TotalDebitoH = FrmReportes.AdoHistorial.Recordset("MDebito")
                                End If
                                If Not IsNull(FrmReportes.AdoHistorial.Recordset("MCredito")) Then
                                     TotalCreditoH = FrmReportes.AdoHistorial.Recordset("MCredito")
                                End If
                                    
                                TotalHistorico = TotalCreditoH - TotalDebitoH
                             End If
                     
                        End If
                      End If
                    End If
                    
                    
                    'encuentra los movimientos que se hicieron de una cuenta entre el rango especificado
                    DoEvents
                    
                    TotalCuenta = 0
                    Total1 = 0
                    FrmReportes.osProgress2.Value = 0
                    Do While Not FrmReportes.DtaConsulta.Recordset.EOF
                      FrmReportes.osProgress2.Visible = True
                        If FrmReportes.osProgress2.Value = 0 Then
                '            FrmReportes.lblProgreso.Caption = "Consultando Registros del periodo seleccionado para la cuenta " & CodigoCuenta
                            FrmReportes.osProgress2.Max = FrmReportes.DtaConsulta.Recordset.RecordCount
                            FrmReportes.osProgress2.Value = FrmReportes.osProgress2.Value + 1
                
                             DoEvents
                        Else
                '            FrmReportes.lblProgreso.Caption = "Consultando Registros del periodo seleccionado para la cuenta " & CodigoCuenta
                            FrmReportes.osProgress2.Value = FrmReportes.osProgress2.Value + 1
                            DoEvents
                        End If
                        
                       TotalDebito = 0
                       TotalCredito = 0
                       TotalDebitoDpto = 0
                       TotalCreditoDpto = 0
                      TipoCuenta = FrmReportes.DtaConsulta.Recordset("TipoCuenta")
                      CodDepartamento = FrmReportes.DtaConsulta.Recordset("VoucherNo")
                      TipoMoneda = FrmReportes.DtaConsulta.Recordset("TipoMoneda")
                      FechaTransaccion = FrmReportes.DtaConsulta.Recordset("FechaTransaccion")
                      NumFecha = FrmReportes.DtaConsulta.Recordset("FechaTransaccion")
                      Fechas = FrmReportes.DtaConsulta.Recordset("FechaTransaccion")
                      FrmReportes.DtaTasas2.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas WHERE (FechaTasas = '" & Format(Fechas, "yyyymmdd") & "')"
                      FrmReportes.DtaTasas2.Refresh
                      If Not FrmReportes.DtaTasas2.Recordset.EOF Then
                        TasaCambio = FrmReportes.DtaTasas2.Recordset("MontoCordobas")
                      Else
                        TasaCambio = 0
                      End If
                     If TasaCambio = 0 Then
                      FrmReportes.osProgress2.Visible = False
                      FrmReportes.osProgress1.Visible = False
                      cadena = "La tasa de Cambio con Fecha: " & Fechas & vbLf
                      cadena = cadena & "no puede ser igual a Cero, el Sistema Contable" & vbLf
                      cadena = cadena & "no contiuara el proceso......"
                      MsgBox cadena, vbCritical, "Sistema Contable"
                      Exit Sub
                     End If
                      If Not IsNull(FrmReportes.DtaConsulta.Recordset("DescripcionCuentas")) Then
                       Descripcion = FrmReportes.DtaConsulta.Recordset("CodCuentas") + "." + FrmReportes.DtaConsulta.Recordset("DescripcionCuentas")
                      Else
                       Descripcion = FrmReportes.DtaConsulta.Recordset("CodCuentas") + "." + "NO TIENE DESCRIPCION????"
                      End If
                      

'                      FrmReportes.AdoConsultas.RecordSource = "SELECT  NTransaccion, FechaTransaccion, VoucherNo, ChequeNo, DescripcionMovimiento, CodCuentas, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, Debito - Credito AS Balance, TCambio, NumeroMovimiento, Beneficiario  From Transacciones  " & _
'                                                              "WHERE (FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') AND (CodCuentas = '" & CodigoCuenta & "') ORDER BY CodCuentas, FechaTransaccion, NTransaccion"

                      If FrmReportes.CmbMoneda.Text = "Córdobas" Then
                        If FrmReportes.ChkQuitarMovimiento.Value = 1 Then
                         FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.NTransaccion, Transacciones.FechaTransaccion, Transacciones.VoucherNo, Transacciones.ChequeNo, Transacciones.DescripcionMovimiento, Transacciones.CodCuentas, Transacciones.TCambio * Transacciones.Debito AS MDebito, Transacciones.TCambio * Transacciones.Credito AS MCredito, Transacciones.Debito - Transacciones.Credito AS Balance, Transacciones.TCambio, Transacciones.NumeroMovimiento, Transacciones.Beneficiario, IndiceTransaccion.Fuente FROM  Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento  " & _
                                                                 "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) AND (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (IndiceTransaccion.Fuente <> 'Cierre') ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"
                        Else
'                        FrmReportes.AdoConsultas.RecordSource = "SELECT  NTransaccion, FechaTransaccion, VoucherNo, ChequeNo, DescripcionMovimiento, CodCuentas, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, Debito - Credito AS Balance, TCambio, NumeroMovimiento, Beneficiario From Transacciones  " & _
'                                                              "WHERE (FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') AND (CodCuentas = '" & CodigoCuenta & "') ORDER BY CodCuentas, FechaTransaccion, NTransaccion "
                         FrmReportes.AdoConsultas.RecordSource = "SELECT  NTransaccion, FechaTransaccion, VoucherNo, ChequeNo, DescripcionMovimiento, CodCuentas, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, Debito - Credito AS Balance, TCambio, NumeroMovimiento, Beneficiario From Transacciones  " & _
                                                              "WHERE (FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') AND (CodCuentas = '" & CodigoCuenta & "') AND (VoucherNo = '" & CodDepartamento & "') ORDER BY CodCuentas, FechaTransaccion, NTransaccion "
                        End If
                      Else
                        FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.NTransaccion, Transacciones.FechaTransaccion, Transacciones.VoucherNo, Transacciones.ChequeNo, Transacciones.DescripcionMovimiento , Transacciones.CodCuentas, Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas) AS MDebito, Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas) AS MCredito, Transacciones.Debito - Transacciones.Credito AS Balance, Transacciones.TCambio, Transacciones.NumeroMovimiento, Transacciones.Beneficiario, Tasas.MontoCordobas FROM  Transacciones INNER JOIN  Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
                                                              "WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') AND (Transacciones.CodCuentas = '" & CodigoCuenta & "') ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"
                      End If
                      
                      
                      
                      
                      FrmReportes.AdoConsultas.Refresh
                      TotalDebito = 0
                      TotalCredito = 0
                      Do While Not FrmReportes.AdoConsultas.Recordset.EOF
                        If Not IsNull(FrmReportes.AdoConsultas.Recordset("MDebito")) Then
                            Debito = FrmReportes.AdoConsultas.Recordset("MDebito")
'                            Debito = TRUNC(Debito, 4)
'                            Debito = Format(Debito, "##,##0.00")
                            TotalDebito = Debito + TotalDebito
                            TotalDebitoDpto = Debito + TotalDebitoDpto
                        End If
                        If Not IsNull(FrmReportes.AdoConsultas.Recordset("MCredito")) Then
                            Credito = FrmReportes.AdoConsultas.Recordset("MCredito")
'                            Credito = TRUNC(Credito, 4)
'                            Credito = Format(Credito, "##,##0.00")
                            TotalCredito = Credito + TotalCredito
                            TotalCreditoDpto = Credito + TotalCreditoDpto
                        End If
                      
                        FrmReportes.AdoConsultas.Recordset.MoveNext
                      Loop
                      
                      Debito = TotalDebito
                      Credito = TotalCredito
                      
'                      If TipoCuenta = "Inventario" Then
'                       cod = 1
'                      End If
'
                                                              
                      If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                
                            If FrmReportes.ChkQuitarMovimiento.Value = 1 Then
                                 FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.NTransaccion, Transacciones.FechaTransaccion, Transacciones.VoucherNo, Transacciones.ChequeNo, Transacciones.DescripcionMovimiento, Transacciones.CodCuentas, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Transacciones.TCambio END AS MDebito, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Transacciones.TCambio END AS MCredito, Transacciones.Debito - Transacciones.Credito AS Balance, Transacciones.TCambio, Transacciones.NumeroMovimiento, Transacciones.Beneficiario, IndiceTransaccion.Nperiodo, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Transacciones.TCambio END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Transacciones.TCambio END AS CreditoD " & _
                                                                         "FROM Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento  " & _
                                                                         "WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') AND (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (IndiceTransaccion.Fuente <> 'Cierre') ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"
'                                FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.NTransaccion, Transacciones.FechaTransaccion, Transacciones.VoucherNo, Transacciones.ChequeNo, Transacciones.DescripcionMovimiento, Transacciones.CodCuentas, Transacciones.TCambio * Transacciones.Debito AS MDebito, Transacciones.TCambio * Transacciones.Credito AS MCredito, Transacciones.Debito - Transacciones.Credito AS Balance, Transacciones.TCambio, Transacciones.NumeroMovimiento, Transacciones.Beneficiario, IndiceTransaccion.Fuente FROM Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
'                                                                        "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) AND (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (IndiceTransaccion.Fuente <> 'Cierre') ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"
                            Else
                                 FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.NTransaccion, Transacciones.FechaTransaccion, Transacciones.VoucherNo, Transacciones.ChequeNo, Transacciones.DescripcionMovimiento, Transacciones.CodCuentas, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Transacciones.TCambio END AS MDebito, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Transacciones.TCambio END AS MCredito, Transacciones.Debito - Transacciones.Credito AS Balance, Transacciones.TCambio, Transacciones.NumeroMovimiento, Transacciones.Beneficiario, IndiceTransaccion.Nperiodo, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Transacciones.TCambio END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Transacciones.TCambio END AS CreditoD " & _
                                                                         "FROM Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento  " & _
                                                                         "WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') AND (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (Transacciones.VoucherNo = '" & CodDepartamento & "') ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"
'                                FrmReportes.AdoConsultas.RecordSource = "SELECT  NTransaccion, FechaTransaccion, VoucherNo, ChequeNo, DescripcionMovimiento, CodCuentas, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, Debito - Credito AS Balance, TCambio, NumeroMovimiento, Beneficiario From Transacciones  " & _
'                                                                        "WHERE (FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') AND (CodCuentas = '" & CodigoCuenta & "') ORDER BY CodCuentas, FechaTransaccion, NTransaccion "
                            End If
                                                                        
                      FrmReportes.AdoConsultas.Refresh
                      TotalDebito = 0
                      TotalCredito = 0
                      Do While Not FrmReportes.AdoConsultas.Recordset.EOF
                        If Not IsNull(FrmReportes.AdoConsultas.Recordset("MDebito")) Then
                          If FrmReportes.CmbMoneda.Text = "Dólares" Then
                            Debito = FrmReportes.AdoConsultas.Recordset("DebitoD")
                          Else
                            Debito = FrmReportes.AdoConsultas.Recordset("MDebito")
                          End If
                            
                          
'                            Debito = TRUNC(Debito, 3)
'                            Debito = Format(Debito, "##,##0.00")
                            TotalDebito = Debito + TotalDebito
                        End If
                        
                        If Not IsNull(FrmReportes.AdoConsultas.Recordset("MCredito")) Then
                          If FrmReportes.CmbMoneda.Text = "Dólares" Then
                            Credito = FrmReportes.AdoConsultas.Recordset("CreditoD")
                          Else
                            Credito = FrmReportes.AdoConsultas.Recordset("MCredito")
                          End If
'                            Credito = TRUNC(Credito, 3)
'                            Credito = Format(Credito, "##,##0.00")
                            TotalCredito = Credito + TotalCredito
                        End If
                      
                        FrmReportes.AdoConsultas.Recordset.MoveNext
                      Loop
                      
                      Debito = TotalDebito
                      Credito = TotalCredito
                        
                      
                        
                        'borrar balanza,  si no funciona, totaldebito y total credito no se usan para balanza, hasta que yo
'                        TotalDebito = TotalDebito + Debito
'                        TotalCredito = TotalCredito + Credito
                        
                        '//////////////////Verifico el tipo de moneda de la cuenta////////////////////
                        If Not TipoMoneda = FrmReportes.CmbMoneda.Text Then
                           Select Case TipoMoneda
                              Case "Córdobas"
                                    TotalCuenta = (Debito - Credito) + TotalCuenta
                          
                              Case "Dólares"
                                    TotalCuenta = (DebitoD - CreditoD) + TotalCuenta
                           
                          End Select
                        Else
                               TotalCuenta = (Debito - Credito) + TotalCuenta
                        End If
                        
                          Total1 = Debito - Credito + Total1
                
'                        Debito = 0
'                        Credito = 0
                      Else
                      
                            If FrmReportes.ChkQuitarMovimiento.Value = 1 Then
                                 FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.NTransaccion, Transacciones.FechaTransaccion, Transacciones.VoucherNo, Transacciones.ChequeNo, Transacciones.DescripcionMovimiento, Transacciones.CodCuentas, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Transacciones.TCambio END AS MDebito, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Transacciones.TCambio END AS MCredito, Transacciones.Debito - Transacciones.Credito AS Balance, Transacciones.TCambio, Transacciones.NumeroMovimiento, Transacciones.Beneficiario, IndiceTransaccion.Nperiodo, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Transacciones.TCambio END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Transacciones.TCambio END AS CreditoD " & _
                                                                         "FROM Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento  " & _
                                                                         "WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') AND (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (IndiceTransaccion.Fuente <> 'Cierre') ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"

                            Else
                                 FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.NTransaccion, Transacciones.FechaTransaccion, Transacciones.VoucherNo, Transacciones.ChequeNo, Transacciones.DescripcionMovimiento, Transacciones.CodCuentas, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Transacciones.TCambio END AS MDebito, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Transacciones.TCambio END AS MCredito, Transacciones.Debito - Transacciones.Credito AS Balance, Transacciones.TCambio, Transacciones.NumeroMovimiento, Transacciones.Beneficiario, IndiceTransaccion.Nperiodo, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Transacciones.TCambio END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Transacciones.TCambio END AS CreditoD " & _
                                                                         "FROM Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento  " & _
                                                                         "WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') AND (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (Transacciones.VoucherNo = '" & CodDepartamento & "') ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"
'                                 FrmReportes.AdoConsultas.RecordSource = "SELECT Transacciones.NTransaccion, Transacciones.FechaTransaccion, Transacciones.VoucherNo, Transacciones.ChequeNo, Transacciones.DescripcionMovimiento, Transacciones.CodCuentas, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Transacciones.TCambio END AS MDebito, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Transacciones.TCambio END AS MCredito, Transacciones.Debito - Transacciones.Credito AS Balance, Transacciones.TCambio, Transacciones.NumeroMovimiento, Transacciones.Beneficiario, IndiceTransaccion.Nperiodo, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Transacciones.TCambio END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Transacciones.TCambio END AS CreditoD " & _
'                                                                         "FROM Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento  " & _
'                                                                         "WHERE (Transacciones.FechaTransaccion BETWEEN '" & Format(FechaIni, "yyyymmdd") & "' AND '" & Format(FechaFin, "yyyymmdd") & "') AND (Transacciones.CodCuentas = '" & CodigoCuenta & "') ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"
                            End If
                                                                        
                      FrmReportes.AdoConsultas.Refresh
                      TotalDebito = 0
                      TotalCredito = 0
                      Do While Not FrmReportes.AdoConsultas.Recordset.EOF
                        If Not IsNull(FrmReportes.AdoConsultas.Recordset("MDebito")) Then
                          If FrmReportes.CmbMoneda.Text = "Dólares" Then
                            Debito = FrmReportes.AdoConsultas.Recordset("DebitoD")
                          Else
                            Debito = FrmReportes.AdoConsultas.Recordset("MDebito")
                          End If
'                            Debito = TRUNC(Debito, 3)
'                            Debito = Format(Debito, "##,##0.00")
                            TotalDebito = Debito + TotalDebito
                        End If
                        If Not IsNull(FrmReportes.AdoConsultas.Recordset("MCredito")) Then
                          If FrmReportes.CmbMoneda.Text = "Dólares" Then
                            Credito = FrmReportes.AdoConsultas.Recordset("CreditoD")
                          Else
                            Credito = FrmReportes.AdoConsultas.Recordset("MCredito")
                          End If
'                            Credito = TRUNC(Credito, 3)
'                            Credito = Format(Credito, "##,##0.00")
                            TotalCredito = Credito + TotalCredito
                        End If
                      
                        FrmReportes.AdoConsultas.Recordset.MoveNext
                      Loop
                      
                      Debito = TotalDebito
                      Credito = TotalCredito
                      
'                         If Not IsNull(FrmReportes.AdoConsultas.Recordset("MDebito")) Then
'                            Debito = FrmReportes.AdoConsultas.Recordset("MDebito")
'                            Debito = TRUNC(Debito, 3)
'                         End If
'                         If Not IsNull(FrmReportes.AdoConsultas.Recordset("MCredito")) Then
'                            Credito = FrmReportes.AdoConsultas.Recordset("MCredito")
'                            Credito = TRUNC(Credito, 3)
'                            Credito = Format(Credito, "##,##0.00")
'                         End If
                         
                        '//////////////////Verifico el tipo de moneda de la cuenta////////////////////
                        If Not TipoMoneda = FrmReportes.CmbMoneda.Text Then
                           Select Case TipoMoneda
                              Case "Córdobas"
                                    TotalCuenta = (Credito - Debito) + TotalCuenta
                          
                              Case "Dólares"
                                    TotalCuenta = (Credito - Debito) + TotalCuenta
                           
                          End Select
                        Else
                               TotalCuenta = (Credito - Debito) + TotalCuenta
                        End If
                         
                         Total1 = Credito - Debito + Total1
'                         Debito = 0
'                         Credito = 0
                      End If
                    
                    
                
                   
                   FrmReportes.DtaConsulta.Recordset.MoveNext
                
                   Loop

         '/////////////////////////////////////////////////////////////////////////////////////////////////////////
         '////////////////////////////////GRABO LAS CUENTAS DEL PERIODO/////////////////////////////////////
         '///////////////////////////////////////////////////////////////////////////////////////////////////////


                    If QUIEN = "Balanza" Then
                '//////////////////////////Busco la Cuenta en Reportes/////////////////////////////////
                       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.KeyGrupoCuenta,Reportes.KeyGrupoSuperior,Reportes.KeyGrupo,Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3,CodCuentas From Reportes Where (((Reportes.KeyGrupo) = '" & CodigoCuenta & "'))"
                       FrmReportes.DtaConsulta2.Refresh
                '       InputBox "", "", FrmReportes.DtaConsulta2.RecordSource
                       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
                
                '          'FrmReportes.DtaConsulta2.Recordset.Edit
                          If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                           If TotalCuenta < 0 Then
                             FrmReportes.DtaConsulta2.Recordset("Haber2") = Credito
                             FrmReportes.DtaConsulta2.Recordset("Debe2") = Debito
'                             Abs(TotalCuenta)
                           Else
                             FrmReportes.DtaConsulta2.Recordset("Debe2") = Debito
                             FrmReportes.DtaConsulta2.Recordset("Haber2") = Credito
'                             TotalCuenta
                           End If
                          Else
                           If TotalCuenta < 0 Then
                               FrmReportes.DtaConsulta2.Recordset("Debe2") = Debito
                               FrmReportes.DtaConsulta2.Recordset("Haber2") = Credito
'                               Abs(TotalCuenta)
                           Else
                               FrmReportes.DtaConsulta2.Recordset("Haber2") = Credito
                               FrmReportes.DtaConsulta2.Recordset("Debe2") = Debito
'                               TotalCuenta
                           End If
                          End If
'                          FrmReportes.DtaConsulta2.Recordset("CodCuentas") = CodigoCuenta
                          FrmReportes.DtaConsulta2.Recordset.Update
                       End If
                       
                       
                       
                '//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
                              Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta"))
                              If Longitud > 1 Then
                                 If Longitud = 5 Then
                                     Nivel = 2
                                 Else
                                     Nivel = (Longitud - 5) / 2
                                     Nivel = Nivel + 2
                                 End If
                              Else
                                    Nivel = 1
                              End If
                       
                '///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
                       Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta")
                       

                       'Nivel = Nivel - 1
                       For i = Nivel To 1 Step -1
                '/////////Busco el Grupo para Sumar los Totaldes
                           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, CodCuentas From Reportes Where (((Reportes.KeyGrupo) = '" & Respuesta & "')) ORDER BY Orden"
                           FrmReportes.DtaConsulta.Refresh
                           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                              FrmReportes.DtaConsulta.Recordset.MoveLast
                '              'FrmReportes.'DtaConsulta.Recordset.Edit
                              If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                                If (TotalCuenta) < 0 Then
                                  FrmReportes.DtaConsulta.Recordset("Haber2") = FrmReportes.DtaConsulta.Recordset("Haber2") + Credito 'Abs(TotalCuenta)
                                  FrmReportes.DtaConsulta.Recordset("Debe2") = FrmReportes.DtaConsulta.Recordset("Debe2") + Debito
                                Else
                                 FrmReportes.DtaConsulta.Recordset("Debe2") = FrmReportes.DtaConsulta.Recordset("Debe2") + Debito 'TotalCuenta
                                 FrmReportes.DtaConsulta.Recordset("Haber2") = FrmReportes.DtaConsulta.Recordset("Haber2") + Credito
                                End If
                              Else
                               If TotalCuenta < 0 Then
                                 FrmReportes.DtaConsulta.Recordset("Debe2") = FrmReportes.DtaConsulta.Recordset("Debe2") + Debito 'Abs(TotalCuenta)
                                 FrmReportes.DtaConsulta.Recordset("Haber2") = FrmReportes.DtaConsulta.Recordset("Haber2") + Credito
                               Else
                                 FrmReportes.DtaConsulta.Recordset("Debe2") = FrmReportes.DtaConsulta.Recordset("Debe2") + Debito
                                 FrmReportes.DtaConsulta.Recordset("Haber2") = FrmReportes.DtaConsulta.Recordset("Haber2") + Credito 'TotalCuenta
                               End If
                              End If
'                              FrmReportes.DtaConsulta.Recordset("CodCuentas") = CodigoCuenta
                              FrmReportes.DtaConsulta.Recordset.Update
                              
                           If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
                              Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
                           End If
                           
                       If Respuesta = "G01000301" Then
                            Cod = 1
                       End If
                
                           End If
                

                       Next
                    
                    
                    
                    ElseIf QUIEN = "BalanzaCodigo" Then
                    
                     '////////////Agrego los Saldos del Periodo Seleccionado////////////////////
                '      FrmReportes.lblProgreso.Caption = "Agregando saldos al Periodo"
                '      FrmReportes.osProgress1.Value = 0
                '      FrmReportes.osProgress1.Max
                
                
                      
                      FrmReportes.DtaReportes.Recordset.AddNew
                      FrmReportes.DtaReportes.Recordset("Descripcion") = Descripcion
                      FrmReportes.DtaReportes.Recordset("CodCuentas") = CodigoCuenta 'Don guillermo
                      If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                        FrmReportes.DtaReportes.Recordset("Debe2") = TotalDebito
                        FrmReportes.DtaReportes.Recordset("Haber2") = TotalCredito
                      Else
                        FrmReportes.DtaReportes.Recordset("Debe2") = TotalDebito
                        FrmReportes.DtaReportes.Recordset("Haber2") = TotalCredito
                      End If
                      FrmReportes.DtaReportes.Recordset!Orden = Orden
                      Orden = Orden + 1
                        
                      FrmReportes.DtaReportes.Recordset.Update
                      
                    
                    ElseIf QUIEN = "Balance" Then
                    
                '//////////////////////////Busco la Cuenta en Reportes/////////////////////////////////
                       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.KeyGrupoCuenta,Reportes.KeyGrupoSuperior,Reportes.KeyGrupo,Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3 From Reportes Where (((Reportes.KeyGrupo) = '" & CodigoCuenta & "'))ORDER BY Orden "
                       FrmReportes.DtaConsulta2.Refresh
                       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
                
                          'FrmReportes.DtaConsulta2.Recordset.Edit
                            FrmReportes.DtaConsulta2.Recordset("Debe1") = TotalCuenta
                          FrmReportes.DtaConsulta2.Recordset.Update
                       End If
                       
                '//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
                              Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta"))
                              If Longitud > 1 Then
                                 If Longitud = 5 Then
                                     Nivel = 2
                                 Else
                                     Nivel = (Longitud - 5) / 2
                                     Nivel = Nivel + 2
                                 End If
                              Else
                                    Nivel = 1
                              End If
                       
                '///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
                       Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta")
                       'Nivel = Nivel - 1
                       For i = Nivel To 1 Step -1
                '/////////Busco el Grupo para Sumar los Totaldes
                           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = '" & Respuesta & "')) ORDER BY Orden"
                '          InputBox "", "", FrmReportes.DtaConsulta.RecordSource
                           FrmReportes.DtaConsulta.Refresh
                           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                              FrmReportes.DtaConsulta.Recordset.MoveLast
                '              'FrmReportes.'DtaConsulta.Recordset.Edit
                                 FrmReportes.DtaConsulta.Recordset("Haber1") = FrmReportes.DtaConsulta.Recordset("Haber1") + TotalCuenta
                              FrmReportes.DtaConsulta.Recordset.Update
                
                           End If
                
                           If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
                              Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
                           End If
                       Next
                '//UTILIDAD
                    ElseIf QUIEN = "UtilidadResultado" Then
                       If TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Then
                         TotalGastos = TotalGastos + TotalCuenta
                    
                       ElseIf TipoCuenta = "Ingresos - Ventas" Then
                         Totalingresos = Totalingresos + TotalCuenta
                       End If
                       
                '       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta From Reportes Where (((Reportes.Descripcion) Like '*Resultado Periodo*'))"
                       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta From Reportes Where (((Reportes.Descripcion) Like '%Resultado Periodo%'))"
                       FrmReportes.DtaConsulta2.Refresh
                       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
                           'FrmReportes.DtaConsulta2.Recordset.Edit
                               FrmReportes.DtaConsulta2.Recordset("Haber1") = Totalingresos - TotalGastos
                           FrmReportes.DtaConsulta2.Recordset.Update
                       End If
                      
                    
                    ElseIf QUIEN = "Utilidad" Then
                       If TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Then
                         TotalGastos = TotalGastos + TotalCuenta
                    
                       ElseIf TipoCuenta = "Ingresos - Ventas" Then
                         Totalingresos = Totalingresos + TotalCuenta
                       End If
                       
                '       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta From Reportes Where (((Reportes.Descripcion) Like '*Resultado Periodo*'))"
                       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta From Reportes Where (((Reportes.Descripcion) Like '%Resultado Periodo%'))"
                       FrmReportes.DtaConsulta2.Refresh
                       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
                           'FrmReportes.DtaConsulta2.Recordset.Edit
                               FrmReportes.DtaConsulta2.Recordset("Debe1") = Totalingresos - TotalGastos
                           FrmReportes.DtaConsulta2.Recordset.Update
                    
                '//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
                              Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupo"))
                              If Longitud > 1 Then
                                 If Longitud = 5 Then
                                     Nivel = 2
                                 Else
                                     Nivel = (Longitud - 5) / 2
                                     Nivel = Nivel + 2
                                 End If
                              Else
                                    Nivel = 1
                              End If
                       
                '///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
                       Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupo")
                       'Nivel = Nivel - 1
                '       For I = Nivel To 1 Step -1
                '/////////Busco el Grupo para Sumar los Totaldes
                           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = 'PC')) ORDER BY Orden"
                           FrmReportes.DtaConsulta.Refresh
                           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                              FrmReportes.DtaConsulta.Recordset.MoveLast
                                 FrmReportes.DtaConsulta.Recordset("Haber1") = (Totalingresos - TotalGastos)
                              FrmReportes.DtaConsulta.Recordset.Update
                
                           End If
                           
                           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = 'C')) ORDER BY Orden"
                           FrmReportes.DtaConsulta.Refresh
                           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                              FrmReportes.DtaConsulta.Recordset.MoveLast
                                 FrmReportes.DtaConsulta.Recordset("Haber1") = (Totalingresos - TotalGastos)
                              FrmReportes.DtaConsulta.Recordset.Update
                
                           End If
                
                '           If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
                '              Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
                '           End If
                '       Next
                      End If
                    ElseIf QUIEN = "Resultado" Then
                '//////////////////////////Busco la Cuenta en Reportes/////////////////////////////////
'                       FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.KeyGrupoCuenta,Reportes.KeyGrupoSuperior,Reportes.KeyGrupo,Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3 From Reportes Where (((Reportes.KeyGrupo) = '" & CodigoCuenta & "'))"
                       FrmReportes.DtaConsulta2.RecordSource = "SELECT  KeyGrupoCuenta, KeyGrupoSuperior, KeyGrupo, Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, CodDepartamento From Reportes WHERE (KeyGrupo = '" & CodigoCuenta & "') AND (CodDepartamento = '" & CodDepartamento & "')"
                       FrmReportes.DtaConsulta2.Refresh
                       If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
                
                          'FrmReportes.DtaConsulta2.Recordset.Edit
                            FrmReportes.DtaConsulta2.Recordset("Debe1") = TotalCuenta
                          FrmReportes.DtaConsulta2.Recordset.Update
                '       End If
                       
                '//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
                              Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta"))
                              If Longitud > 1 Then
                                 If Longitud = 5 Then
                                     Nivel = 2
                                 Else
                                     Nivel = (Longitud - 5) / 2
                                     Nivel = Nivel + 2
                                 End If
                              Else
                                    Nivel = 1
                              End If
                       
                '///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
                       Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta")
                       Else
                       
                       MsgBox "La cuenta Tiene Saldo y no aparece en la Estructura, Cuenta: " & CodigoCuenta, vbCritical
                        
                       End If
                       'Nivel = Nivel - 1
                       For i = Nivel To 1 Step -1
                '/////////Busco el Grupo para Sumar los Totaldes
                           FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = '" & Respuesta & "')) ORDER BY Orden"
                '           InputBox "", "", DtaConsulta.RecordSource
                           FrmReportes.DtaConsulta.Refresh
                           If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                              FrmReportes.DtaConsulta.Recordset.MoveLast
                              'FrmReportes.'DtaConsulta.Recordset.Edit
                                 FrmReportes.DtaConsulta.Recordset("Haber1") = FrmReportes.DtaConsulta.Recordset("Haber1") + TotalCuenta
                              FrmReportes.DtaConsulta.Recordset.Update
                
                           End If
                
                           If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
                              Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
                           End If
                       Next
                    End If
                    
                    
                    '---------------------------------------------------------------------------------------------------
                    '-----------------BUSCO DEPARTAMENTO PARA CARGAR LOS SALDOS ---------------------------------------
                    '---------------------------------------------------------------------------------------------------
                     FrmReportes.DtaConsulta.RecordSource = "SELECT  KeyGrupoCuenta, KeyGrupoSuperior, KeyGrupo, Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, CodDepartamento From Reportes WHERE (KeyGrupo = '" & CodigoGrupo & "') AND (CodDepartamento = '" & CodDepartamento & "') AND (Descripcion LIKE '%Total%')"
                    FrmReportes.DtaConsulta.Refresh
                    If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                          FrmReportes.DtaConsulta.Recordset("Haber1") = FrmReportes.DtaConsulta.Recordset("Haber1") + TotalCuenta
                          FrmReportes.DtaConsulta.Recordset.Update
                    End If
                    
                
                   FrmReportes.DtaHistorial.Recordset.MoveNext
                 
                  Loop
  
                 FrmReportes.osProgress2.Visible = False
                 
  
'*************************************************************************************************************
'*************************************************************************************************************
'*************************************************************************************************************
        '/////////////////////////////////////////////////////////////////////////////////////////////////////
        '//////////////////PERIODO ANTERIOR///////////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////////////////////
'*************************************************************************************************************
'*************************************************************************************************************
'*************************************************************************************************************

        If FrmReportes.CmbReportes.Text = "LISTA CUENTAS X COBRAR" Or FrmReportes.CmbReportes.Text = "LISTA CUENTAS X PAGAR" Then
          QUIEN = "SaldoCuentas"
          
        End If
  
        If FrmReportes.ChkBalanza.Value = False Then
          
            Debito = 0
            Credito = 0
            Totalingresos = 0
            TotalGastos = 0
            Sql = ""
            'Busco que cuentas tienen saldo
            If QUIEN = "Utilidad" Then
                Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) <'" & Format(FechaIni, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda" & vbLf
                Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Ingresos - Ventas' Or (Cuentas.TipoCuenta) = 'Costos' Or (Cuentas.TipoCuenta) = 'Gastos')) ORDER BY Cuentas.CodCuentas"
            ElseIf QUIEN = "BalanzaCodigo" Then
'                SQL = "SELECT Cuentas.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio,5)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito,5)) AS MCredito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio,5)) - SUM(ROUND(Transacciones.Credito * Transacciones.TCambio,5)) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM  Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas " & _
'                      "WHERE (Transacciones.FechaTransaccion < '" & Format(FechaIni, "yyyymmdd") & "') GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda Having (SUM(Round(Transacciones.Debito * Transacciones.TCambio,5)) - SUM(Round(Transacciones.Credito * Transacciones.TCambio,5)) <> 0) ORDER BY Cuentas.CodCuentas"
                 Sql = "SELECT Transacciones.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 2)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito, 2)) AS MCredito, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas  " & _
                       "WHERE (Transacciones.FechaTransaccion < '" & Format(FechaIni, "yyyymmdd") & "') GROUP BY Transacciones.CodCuentas HAVING  (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "')"
            ElseIf QUIEN = "SaldoCuentas" Then
                 QUIEN = "BalanzaCodigo"
                 Sql = "SELECT Transacciones.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 2)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito, 2)) AS MCredito, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas  " & _
                       "WHERE (Transacciones.FechaTransaccion < '" & Format(FechaIni, "yyyymmdd") & "') GROUP BY Transacciones.CodCuentas HAVING  (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') ORDER BY Transacciones.CodCuentas"  'AND (MAX(Cuentas.TipoCuenta) = 'Cuentas x Cobrar')
            
            ElseIf QUIEN = "Balanza" Then
'                SQl = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) <'" & Format(FechaIni, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda" & vbLf
'                SQl = SQl & "ORDER BY Cuentas.CodCuentas"
        
                Sql = "SELECT Cuentas.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio) - SUM(Transacciones.Credito * Transacciones.TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, MAX(Cuentas.KeyGrupo) AS KeyGrupo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (Transacciones.FechaTransaccion < '" & Format(FechaIni, "yyyymmdd") & "') GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda HAVING (MAX(Cuentas.KeyGrupo) BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') ORDER BY Cuentas.CodCuentas"
            ElseIf QUIEN = "Balance" Then
              
                
                Sql = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) <'" & Format(FechaIni, "yyyymmdd") & "')) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda" & vbLf
                Sql = Sql & "Having (((Cuentas.TipoCuenta) = 'Otros Activos' Or (Cuentas.TipoCuenta) = 'Caja' Or (Cuentas.TipoCuenta) = 'Bancos' Or (Cuentas.TipoCuenta) = 'Cuentas x Cobrar' Or (Cuentas.TipoCuenta) = 'Inventario' Or (Cuentas.TipoCuenta) = 'Papeleria - Utiles' Or (Cuentas.TipoCuenta) = 'Activo Fijo' Or (Cuentas.TipoCuenta) = 'Otros Pasivos' Or (Cuentas.TipoCuenta) = 'Cuentas x Pagar' Or (Cuentas.TipoCuenta) = 'Pasivo' Or (Cuentas.TipoCuenta) = 'Capital')) ORDER BY Cuentas.CodCuentas"
            End If
            
              If Sql <> "" Then
                FrmReportes.DtaHistorial.RecordSource = Sql
                FrmReportes.DtaHistorial.Refresh
              End If
        
                FrmReportes.LblProgreso.Caption = "Consultando Registros del Periodo Anterior para " & QUIEN
                FrmReportes.osProgress1.Value = 0
            
            If Not FrmReportes.DtaHistorial.Recordset.EOF Then
                FrmReportes.osProgress1.Max = FrmReportes.DtaHistorial.Recordset.RecordCount
                FrmReportes.DtaHistorial.Refresh
            Else
                ' Exit Sub
            End If
        
            Do While Not FrmReportes.DtaHistorial.Recordset.EOF
                '////////Consulto los registros del periodo ANTERIOR.///////////
                FrmReportes.osProgress1.Value = FrmReportes.osProgress1.Value + 1
                CodigoCuenta = FrmReportes.DtaHistorial.Recordset("CodCuentas")
                


                
                FrmReportes.LblProgreso.Caption = "Consultando Registros del Periodo Anterior para la Cuenta " & CodigoCuenta
                DoEvents
            
                FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, Transacciones.FechaTransaccion, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio - Transacciones.TCambio * Transacciones.Credito) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Tasas.MontoCordobas) AS MontoCordobas, MAX(Tasas.MontoLibras) AS MontoLibras, MAX(Transacciones.NTransaccion) AS NTransaccion FROM  Tasas INNER JOIN  Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Tasas.FechaTasas = Transacciones.FechaTasas GROUP BY Cuentas.CodCuentas, Transacciones.FechaTransaccion  " & _
                                                       "HAVING (Cuentas.CodCuentas = '" & CodigoCuenta & "') AND (Transacciones.FechaTransaccion < CONVERT(DATETIME,'" & Format(FechaIni, "yyyymmdd") & "', 102)) ORDER BY Cuentas.CodCuentas"

'                FrmReportes.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 2)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito, 2)) AS MCredito, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas  " & _
'                                                       "WHERE (Transacciones.FechaTransaccion < '" & Format(FechaIni, "yyyymmdd") & "') GROUP BY Transacciones.CodCuentas HAVING  (Transacciones.CodCuentas = '" & CodigoCuenta & "')"
                FrmReportes.DtaConsulta.Refresh
                
                TotalCuenta = 0
                Total1 = 0
                FrmReportes.osProgress2.Value = 0
            
                Do While Not FrmReportes.DtaConsulta.Recordset.EOF
                    FrmReportes.osProgress2.Visible = True
                    If FrmReportes.osProgress2.Value = 0 Then
                '            FrmReportes.lblProgreso.Caption = "Consultando Registros del periodo seleccionado para la cuenta " & CodigoCuenta
                        FrmReportes.osProgress2.Max = FrmReportes.DtaConsulta.Recordset.RecordCount
                        FrmReportes.osProgress2.Value = FrmReportes.osProgress2.Value + 1
        
                        DoEvents
                    Else
        '               FrmReportes.lblProgreso.Caption = "Consultando Registros del periodo seleccionado para la cuenta " & CodigoCuenta
                        FrmReportes.osProgress2.Value = FrmReportes.osProgress2.Value + 1
                        DoEvents
                    End If
        
                    TotalDebito = 0
                    TotalCredito = 0
                    TipoCuenta = FrmReportes.DtaConsulta.Recordset("TipoCuenta")
                    TipoMoneda = FrmReportes.DtaConsulta.Recordset("TipoMoneda")
                    FechaTransaccion = FrmReportes.DtaConsulta.Recordset("FechaTransaccion")
                    Fechas1 = FrmReportes.DtaConsulta.Recordset("FechaTransaccion")
                    TasaCambio = FrmReportes.DtaConsulta.Recordset("MontoCordobas")
                    If TasaCambio = 0 Then
                        cadena = "La tasa de Cambio de Cambio con Fecha: " & Fechas1 & vbLf
                        cadena = cadena & "no puede ser igual a Cero, el Sistema Contable" & vbLf
                        cadena = cadena & "no contiuara el proceso......"
                        MsgBox cadena, vbCritical, "Sistema Contable"
                        Exit Sub
                    End If
                    If Not IsNull(FrmReportes.DtaConsulta.Recordset("DescripcionCuentas")) Then
                     Descripcion = FrmReportes.DtaConsulta.Recordset("CodCuentas") + "." + FrmReportes.DtaConsulta.Recordset("DescripcionCuentas")
                    Else
                     Descripcion = FrmReportes.DtaConsulta.Recordset("CodCuentas") + "." + "NO TIEN DESCRIPCION???"
                    End If
                    
'                      FrmReportes.AdoConsultas.RecordSource = "SELECT CodCuentas, SUM(ROUND(Debito * TCambio, 2)) AS MDebito, SUM(ROUND(TCambio * Credito, 2)) AS MCredito From Transacciones  " & _
'                                                              "WHERE (FechaTransaccion < '" & Format(FechaIni, "yyyy-mm-dd") & "') GROUP BY CodCuentas HAVING (CodCuentas = '" & CodigoCuenta & "')"
'                      FrmReportes.AdoConsultas.Refresh
                    
                    
                    If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                     If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                         If Not IsNull(FrmReportes.DtaConsulta.Recordset("MDebito")) Then
                             Debito = FrmReportes.DtaConsulta.Recordset("MDebito")
'                             Debito = TRUNC(Debito, 5)
'                             Debito = Format(Debito, "##,##0.000")
                        End If
                        If Not IsNull(FrmReportes.DtaConsulta.Recordset("MCredito")) Then
                             Credito = FrmReportes.DtaConsulta.Recordset("MCredito")
'                             Credito = TRUNC(Credito, 5)
'                             Credito = Format(Credito, "##,##0.000")
                         End If
                     End If
                        Total1 = Debito - Credito + Total1
        
                '//////////////////Verifico el tipo de moneda de la cuenta////////////////////
                If Not TipoMoneda = FrmReportes.CmbMoneda.Text Then
                   Select Case TipoMoneda
                      Case "Córdobas"
                            TotalCuenta = (Debito - Credito) / TasaCambio + TotalCuenta
        
                      Case "Dólares"
                            TotalCuenta = (Debito - Credito) * TasaCambio + TotalCuenta
        
                  End Select
                Else
                       TotalCuenta = (Debito - Credito) + TotalCuenta
                End If
        
                Debito = 0
                Credito = 0
              Else
               If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                 If Not IsNull(FrmReportes.DtaConsulta.Recordset("MDebito")) Then
                    Debito = FrmReportes.DtaConsulta.Recordset("MDebito")
                 End If
                 If Not IsNull(FrmReportes.DtaConsulta.Recordset("MCredito")) Then
                    Credito = FrmReportes.DtaConsulta.Recordset("MCredito")
                 End If
               End If
        
                '//////////////////Verifico el tipo de moneda de la cuenta////////////////////
                If Not TipoMoneda = FrmReportes.CmbMoneda.Text Then
                   Select Case TipoMoneda
                      Case "Córdobas"
                            TotalCuenta = (Credito - Debito) / TasaCambio + TotalCuenta
        
                      Case "Dólares"
                            TotalCuenta = (Credito - Debito) * TasaCambio + TotalCuenta
        
                  End Select
                Else
                       TotalCuenta = (Credito - Debito) + TotalCuenta
                End If
        
                 Total1 = Credito - Debito + Total1
                 Debito = 0
                 Credito = 0
              End If
        
            FrmReportes.DtaConsulta.Recordset.MoveNext
        
           Loop
           
        'End If
        
        
           
            '//////////////////////////////////////////////////////////////////////
            '////////////////////GRABO LOS REGISTROS DEL PERIODO ANTERIOR//////////
            '//////////////////EN LA TABLA REPORTES////////////////////////////////
            If QUIEN = "Balanza" Then
            '//////////////////////////Busco la Cuenta en Reportes/////////////////////////////////
               FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.KeyGrupoCuenta,Reportes.KeyGrupoSuperior,Reportes.KeyGrupo,Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3,reportes.orden From Reportes Where (((Reportes.KeyGrupo) = '" & CodigoCuenta & "'))"
               FrmReportes.DtaConsulta2.Refresh
               If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
        
                  'FrmReportes.DtaConsulta2.Recordset.Edit
                  If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                    If TotalCuenta < 0 Then
                     FrmReportes.DtaConsulta2.Recordset("Haber1") = Abs(TotalCuenta)
                    Else
                     FrmReportes.DtaConsulta2.Recordset("Debe1") = TotalCuenta
                    End If
                  Else
                    If TotalCuenta < 0 Then
                      FrmReportes.DtaConsulta2.Recordset("Debe1") = Abs(TotalCuenta)
                    Else
                     FrmReportes.DtaConsulta2.Recordset("Haber1") = TotalCuenta
                    End If
                  End If
        '          FrmReportes.DtaConsulta2.Recordset!Orden = Orden
                  Orden = Orden + 1
                  FrmReportes.DtaConsulta2.Recordset.Update
               End If
               
                '//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
                      Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta"))
                      If Longitud > 1 Then
                         If Longitud = 5 Then
                             Nivel = 2
                         Else
                             Nivel = (Longitud - 5) / 2
                             Nivel = Nivel + 2
                         End If
                      Else
                            Nivel = 1
                      End If
               
                '///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
               Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta")
               'Nivel = Nivel - 1
               For i = Nivel To 1 Step -1
                '/////////Busco el Grupo para Sumar los Totaldes
                   FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior,reportes.orden From Reportes Where (((Reportes.KeyGrupo) = '" & Respuesta & "')) ORDER BY Orden"
                   FrmReportes.DtaConsulta.Refresh
                   If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                      FrmReportes.DtaConsulta.Recordset.MoveLast
                      'FrmReportes.'DtaConsulta.Recordset.Edit
                      If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                        If TotalCuenta < 0 Then
                          FrmReportes.DtaConsulta.Recordset("Haber1") = FrmReportes.DtaConsulta.Recordset("Haber1") + Abs(TotalCuenta)
                        Else
                         FrmReportes.DtaConsulta.Recordset("Debe1") = FrmReportes.DtaConsulta.Recordset("Debe1") + TotalCuenta
                        End If
                      Else
                        If TotalCuenta < 0 Then
                         FrmReportes.DtaConsulta.Recordset("Debe1") = FrmReportes.DtaConsulta.Recordset("Debe1") + Abs(TotalCuenta)
                        Else
                         FrmReportes.DtaConsulta.Recordset("Haber1") = FrmReportes.DtaConsulta.Recordset("Haber1") + TotalCuenta
                        End If
                      End If
                      If 1 = 0 Then FrmReportes.DtaConsulta.Recordset!Orden = Orden + 1
                      Orden = Orden + 1
                      
                      FrmReportes.DtaConsulta.Recordset.Update
                      
                   If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
                      Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
                   End If
        
                   End If
        

               Next
            
            
            
            ElseIf QUIEN = "BalanzaCodigo" Then
            
              '////////////Agrego los Saldos del PeriodoAnterior////////////////////
              
              Descripcion = Replace(Descripcion, "'", "")
              
               FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3,reportes.orden,Reportes.CodCuentas From Reportes Where (((Reportes.Descripcion) = '" & Descripcion & "'))"
               FrmReportes.DtaConsulta2.Refresh
               If FrmReportes.DtaConsulta2.Recordset.EOF Then
                 FrmReportes.DtaConsulta2.Recordset.AddNew
                    FrmReportes.DtaConsulta2.Recordset("Descripcion") = Descripcion
                    FrmReportes.DtaConsulta2.Recordset("CodCuentas") = CodigoCuenta
                     FrmReportes.DtaConsulta2.Recordset!Orden = Orden
                     Orden = Orden + 1
                   Else
                     'FrmReportes.DtaConsulta2.Recordset.Edit
                   End If
                  
                   If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                     If TotalCuenta < 0 Then
                       FrmReportes.DtaConsulta2.Recordset("Haber1") = Abs(TotalCuenta)
                     Else
                      FrmReportes.DtaConsulta2.Recordset("Debe1") = TotalCuenta
                     End If
                   Else
                     If TotalCuenta < 0 Then
                      FrmReportes.DtaConsulta2.Recordset("Debe1") = Abs(TotalCuenta)
                     Else
                      FrmReportes.DtaConsulta2.Recordset("Haber1") = TotalCuenta
                     End If
                   End If
            '       FrmReportes.DtaConsulta2.Recordset!Orden = Orden
            '       Orden = Orden + 1
                   FrmReportes.DtaConsulta2.Recordset.Update
                   
               
               '////////////////////AGREGO LOS SALDOS ANTERIORES DEL BALANCE///////////
               ElseIf QUIEN = "Balance" Then
            '//////////////////////////Busco la Cuenta en Reportes/////////////////////////////////
                   FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.KeyGrupoCuenta,Reportes.KeyGrupoSuperior,Reportes.KeyGrupo,Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3 From Reportes Where (((Reportes.KeyGrupo) = '" & CodigoCuenta & "'))ORDER BY Orden "
                   FrmReportes.DtaConsulta2.Refresh
                   If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
            
                      'FrmReportes.DtaConsulta2.Recordset.Edit
                        FrmReportes.DtaConsulta2.Recordset("Debe1") = TotalCuenta + FrmReportes.DtaConsulta2.Recordset("Debe1")
                      FrmReportes.DtaConsulta2.Recordset.Update
                   End If
            
            '//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
                          Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta"))
                          If Longitud > 1 Then
                             If Longitud = 5 Then
                                 Nivel = 2
                             Else
                                 Nivel = (Longitud - 5) / 2
                                 Nivel = Nivel + 2
                             End If
                          Else
                                Nivel = 1
                          End If
            
            '///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
                   Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupoCuenta")
                   'Nivel = Nivel - 1
                   For i = Nivel To 1 Step -1
            '/////////Busco el Grupo para Sumar los Totaldes
                       FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = '" & Respuesta & "')) ORDER BY Orden"
            '          InputBox "", "", FrmReportes.DtaConsulta.RecordSource
                       FrmReportes.DtaConsulta.Refresh
                       If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                          FrmReportes.DtaConsulta.Recordset.MoveLast
            '              'FrmReportes.'DtaConsulta.Recordset.Edit
                             FrmReportes.DtaConsulta.Recordset("Haber1") = FrmReportes.DtaConsulta.Recordset("Haber1") + TotalCuenta
                          FrmReportes.DtaConsulta.Recordset.Update
            
                       End If
            
                       If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
                          Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
                       End If
                   Next
               
             
             '//////////////////AGREGO LA UTILIDAD AL BALANCE//////////////////////////
             ElseIf QUIEN = "Utilidad" Then
                   If TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Then
                     TotalGastos = TotalCuenta
                     Totalingresos = 0
                   ElseIf TipoCuenta = "Ingresos - Ventas" Then
                     Totalingresos = TotalCuenta
                     TotalGastos = 0
                   End If
                   
                   FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior, Reportes.KeyGrupoCuenta From Reportes Where (((Reportes.Descripcion) Like '%Resultado Periodo%'))"
                   FrmReportes.DtaConsulta2.Refresh
                   If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
                           FrmReportes.DtaConsulta2.Recordset("Debe1") = Totalingresos - TotalGastos + FrmReportes.DtaConsulta2.Recordset("Debe1")
                       FrmReportes.DtaConsulta2.Recordset.Update
                
            '//////////////////////IDENTIFICO EL NIVEL DEL PADRE////////////////////////////////////////////////////
                          Longitud = Len(FrmReportes.DtaConsulta2.Recordset("KeyGrupo"))
                          If Longitud > 1 Then
                             If Longitud = 5 Then
                                 Nivel = 2
                             Else
                                 Nivel = (Longitud - 5) / 2
                                 Nivel = Nivel + 2
                             End If
                          Else
                                Nivel = 1
                          End If
                   
            '///////////////////////////Ahora le Sumo el Saldo a los Grupos Superiores/////////////////
                   Respuesta = FrmReportes.DtaConsulta2.Recordset("KeyGrupo")
                   'Nivel = Nivel - 1
            '       For I = Nivel To 1 Step -1
            '/////////Busco el Grupo para Sumar los Totaldes
                       FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = 'PC')) ORDER BY Orden"
                       FrmReportes.DtaConsulta.Refresh
                       If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                          FrmReportes.DtaConsulta.Recordset.MoveLast
                             FrmReportes.DtaConsulta.Recordset("Haber1") = (Totalingresos - TotalGastos) + FrmReportes.DtaConsulta.Recordset("Haber1")
                          FrmReportes.DtaConsulta.Recordset.Update
            
                       End If
                       
                       FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.KeyGrupo, Reportes.KeyGrupoSuperior From Reportes Where (((Reportes.KeyGrupo) = 'C')) ORDER BY Orden"
                       FrmReportes.DtaConsulta.Refresh
                       If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                          FrmReportes.DtaConsulta.Recordset.MoveLast
                             FrmReportes.DtaConsulta.Recordset("Haber1") = (Totalingresos - TotalGastos) + FrmReportes.DtaConsulta.Recordset("Haber1")
                          FrmReportes.DtaConsulta.Recordset.Update
            
                       End If
            
            '           If Not IsNull(FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")) Then
            '              Respuesta = FrmReportes.DtaConsulta.Recordset("KeyGrupoSuperior")
            '           End If
            '       Next
                  End If
              
             
             
                ElseIf QUIEN = "Resultado" Then
            
                End If
            
               
               FrmReportes.DtaHistorial.Recordset.MoveNext
              Loop
              
            End If '////////FIN DEL IF PARA If FrmReportes.ChkBalanza.Value = True Then
            
            
              
            '/////////////////////////////////////////////////////////////////////////////////////////
            '//////////////////////SUMO LOS TOTALES DEL PASIVO + CAPITAL/////////////////////////////
            '/////////////////////////////////////////////////////////////////////////////////////////
              
            If QUIEN = "Balance" Then
                  FrmReportes.DtaConsulta2.RecordSource = "SELECT Sum(Reportes.Debe1) AS SumaDeDebe1, Sum(Reportes.Haber1) AS SumaDeHaber1 From Reportes Where (((Reportes.Descripcion) Like 'Total%') And ((Reportes.KeyGrupo) = 'B' Or (Reportes.KeyGrupo) = 'C'))"
            
               FrmReportes.DtaConsulta2.Refresh
               If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
                  FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.KeyGrupo From Reportes Where (((Reportes.KeyGrupo) = 'PC'))"
                  FrmReportes.DtaConsulta.Refresh
                  If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                   If Not IsNull(FrmReportes.DtaConsulta2.Recordset("SumaDeHaber1")) Then
                    TotalCuenta = FrmReportes.DtaConsulta2.Recordset("SumaDeHaber1")
                    'FrmReportes.'DtaConsulta.Recordset.Edit
                     FrmReportes.DtaConsulta.Recordset("Haber1") = TotalCuenta
                    FrmReportes.DtaConsulta.Recordset.Update
                   End If
                  End If
               End If
               
             ElseIf QUIEN = "Resultado" Then
                FrmReportes.DtaConsulta2.RecordSource = "SELECT Sum(Reportes.Debe1) AS SumaDeDebe1, Sum(Reportes.Haber1) AS SumaDeHaber1 From Reportes Where (((Reportes.Descripcion) Like 'Total%') And ((Reportes.KeyGrupo) = 'G' Or (Reportes.KeyGrupo) = 'O'))"
'               FrmReportes.DtaConsulta2.RecordSource = "SELECT Sum(Reportes.Debe1) AS SumaDeDebe1, Sum(Reportes.Haber1) AS SumaDeHaber1 From Reportes Where (((Reportes.KeyGrupo) = 'G' Or (Reportes.KeyGrupo) = 'O'))"
               FrmReportes.DtaConsulta2.Refresh
               If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
                  FrmReportes.DtaConsulta.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.KeyGrupo From Reportes Where (((Reportes.KeyGrupo) = 'CG'))"
                  FrmReportes.DtaConsulta.Refresh
                  If Not FrmReportes.DtaConsulta.Recordset.EOF Then
                   If Not IsNull(FrmReportes.DtaConsulta2.Recordset("SumaDeHaber1")) Then
                    TotalCuenta = FrmReportes.DtaConsulta2.Recordset("SumaDeHaber1")
                    'FrmReportes.'DtaConsulta.Recordset.Edit
                     FrmReportes.DtaConsulta.Recordset("Haber1") = TotalCuenta
                    FrmReportes.DtaConsulta.Recordset.Update
                   End If
                  End If
               End If
             End If
  
     '////////////Agrego los Saldos de los acumulados////////////////////
'   FrmReportes.DtaConsulta2.RecordSource = "SELECT Reportes.Descripcion, Reportes.Debe1, Reportes.Haber1, Reportes.Debe2, Reportes.Haber2, Reportes.Debe3, Reportes.Haber3, Reportes.Debe1+Reportes.Debe2 AS TotalDebe, Reportes.Haber1+Reportes.Haber2 AS TotalHaber From Reportes"
'   FrmReportes.DtaConsulta2.Refresh
'
'   FrmReportes.LblProgreso.Caption = "Agregando Saldos Acumulados a las Cuentas"
'    FrmReportes.osProgress1.Value = 0
'    FrmReportes.osProgress1.Max = FrmReportes.DtaConsulta2.Recordset.RecordCount
'
'  Do While Not FrmReportes.DtaConsulta2.Recordset.EOF
'      FrmReportes.osProgress1.Value = FrmReportes.osProgress1.Value + 1
'      FrmReportes.DtaConsulta2.Recordset("Debe3") = FrmReportes.DtaConsulta2.Recordset("TotalDebe")
'      FrmReportes.DtaConsulta2.Recordset("Haber3") = FrmReportes.DtaConsulta2.Recordset("TotalHaber")
'
'    FrmReportes.DtaConsulta2.Recordset.Update
'   FrmReportes.DtaConsulta2.Recordset.MoveNext
'   Loop
'
'   FrmReportes.DtaConsulta2.Refresh
    
Ejecutar.Execute "Update Reportes set debe3=debe1+debe2,haber3=haber1+haber2"
 
End Sub





Public Sub LeeTecla()
Select Case Lectura
       
       Case 42
          Lectura = "%"
       Case 48
         Lectura = "0"
       Case 49
         Lectura = "1"
       Case 50
           Lectura = "2"
       Case 51
            Lectura = "3"
       Case 52
            Lectura = 4
       Case 53
             Lectura = 5
        Case 54
             Lectura = 6
        Case 55
             Lectura = 7
        Case 56
             Lectura = 8
        Case 57
              Lectura = 9
        Case 59
              Lectura = "Ñ"
        Case 47
              Lectura = "/"
        Case 81
             Lectura = "Q"
        Case 77
             Lectura = "M"
         Case 87
             Lectura = "W"
         Case 69
             Lectura = "E"
         Case 82
             Lectura = "R"
        Case 84
              Lectura = "T"
        Case 89
              Lectura = "Y"
        Case 85
              Lectura = "U"
                
         Case 112
             Lectura = "P"
         Case 113
             Lectura = "Q"
         Case 119
             Lectura = "W"
         Case 101
             Lectura = "E"
         Case 114
             Lectura = "R"
         Case 116
             Lectura = "T"
         Case 121
             Lectura = "Y"
         Case 117
             Lectura = "U"
         Case 105
             Lectura = "I"
         Case 111
             Lectura = "O"
         Case 97
             Lectura = "A"
         Case 115
             Lectura = "S"
         Case 100
             Lectura = "D"
         Case 102
             Lectura = "F"
         Case 103
             Lectura = "G"
         Case 104
             Lectura = "H"
         Case 106
             Lectura = "J"
         Case 107
             Lectura = "K"
         Case 108
             Lectura = "L"
         Case 122
             Lectura = "Z"
         Case 120
             Lectura = "X"
         Case 99
             Lectura = "C"
         Case 118
             Lectura = "V"
         Case 98
             Lectura = "B"
         Case 110
             Lectura = "N"
         Case 109
             Lectura = "M"
        Case 73
             Lectura = "I"
         Case 79
             Lectura = "O"
         Case 80
             Lectura = "P"
         Case 65
             Lectura = "A"
        Case 83
              Lectura = "S"
        Case 68
              Lectura = "D"
        Case 70
              Lectura = "F"
        Case 71
             Lectura = "G"
         Case 72
             Lectura = "H"
         Case 74
             Lectura = "J"
         Case 75
             Lectura = "K"
         Case 76
             Lectura = "L"
         Case 90
             Lectura = "Z"
        Case 88
              Lectura = "X"
        Case 67
              Lectura = "C"
        Case 86
              Lectura = "V"
        Case 66
             Lectura = "B"
         Case 78
             Lectura = "N"
         Case 77
             Lectura = "M"
        Case 32
             Lectura = " "
        Case 189
             Lectura = "/"
        Case 188
             Lectura = ","
        Case 190
             Lectura = "."
             
             
End Select
End Sub

Public Sub Periodos(FechaCierre As Date)
  Dim mes As Integer
  mes = Month(FechaCierre)
  






End Sub




Public Sub ControlErrores()
     Dim cadena As String
     Select Case err
        Case 53
             cadena = cadena & "No se ha encontrado el Archivo" & vbLf
             cadena = cadena & "que se espicifico inicialmente " & vbLf
             MsgBox Prompt:=cadena, Buttons:=vbExclamation, Title:="Sistema Contable"
             cadena = ""
        Case 3464
             cadena = cadena & "A ocurrido un Error de Criterios" & vbLf
             cadena = cadena & "En el Tipo de Datos, No Coinciden " & vbLf
             cadena = cadena & "Los Tipos con la Expresion de Criterios."
             MsgBox Prompt:=cadena, Buttons:=vbExclamation, Title:="Sistema Contable"
             cadena = ""
        Case 6
             cadena = cadena & "Desbordamiento de Datos en el sistema" & vbLf
             cadena = cadena & "Este error se debe por que  digito muchos Datos " & vbLf
             cadena = cadena & "Consulte su Soporte Tecnico 07788301"
             MsgBox Prompt:=cadena, Buttons:=vbExclamation, Title:="Sistema Contable"
             cadena = ""
        Case 94
             cadena = cadena & "El Registro tiene un Datos que es Nulo" & vbLf
             cadena = cadena & "Debe corregir este error desde las Base " & vbLf
             cadena = cadena & "de Datos, Consulte su Soporte Tecnico 07788301"
             MsgBox Prompt:=cadena, Buttons:=vbExclamation, Title:="Sistema Contable"
             cadena = ""
        Case 76
             cadena = cadena & "La ruta de la Base de Datos no esta" & vbLf
             cadena = cadena & "en el Disco Duro, Esto puede causar " & vbLf
             cadena = cadena & "Conflicto, la Ruta es: C:\facturacion"
             MsgBox Prompt:=cadena, Buttons:=vbExclamation, Title:="Zeus  Facturacion"
             cadena = ""
        Case 3022
             cadena = cadena & "Se estan Creando PK Dupliacados, por Fabor" & vbLf
             cadena = cadena & "Si esta en facturacion, verifique el " & vbLf
             cadena = cadena & "el numero consecutivo."
             MsgBox cadena, vbInformation, "Error 3022:Zeus Facturacion"
        Case 13
             MsgBox "No se Continuara el Proceso", vbInformation, "Error 13:Zeus Facturacion"
                          
        Case 484
             MsgBox "Controlador de la Impresora no Disponible WIN.INI", vbInformation, "Error de Impresion 484:Zeus Facturacion"
        Case 483
             MsgBox "El controlador de la Impresora no admite esta Propiedad", vbInformation, "Error de Impresion 483:Zeus Facturacion"
        Case 482
             MsgBox "Error de la Impresora", vbInformation, "Error de Impresion 482:Sistema de Facturacion"
        Case 396
              MsgBox "Imposible establecer la Propiedad dentro de la Pag.", vbInformation, "Error de Impresion 396:Zeus Facturacion"
        Case 91
              MsgBox "Error grabe ocurrido con la Base de Datos", vbInformation, "Error 91:Zeus Facturacion"
        Case 424
             MsgBox "No se Ha encontrado el Objeto Asociado", vbInformation, "Error 424:Zeus Facturacion"
        Case 53
             MsgBox "El Archivo o la Foto no se Ha encontrado", vbInformation, "Error 53:Zeus Facturacion"
        Case 364
             MsgBox "No se Puede abrir esta Ventana", vbCritical, "Error de Registro 364:Zeus Facturacion"
        Case 380
             cadena = cadena & "       El tipo de Dato No es Correcto,Verifique" & vbLf
             cadena = cadena & "   Por fabor la Configuracion Regional de Windows" & vbLf
             cadena = cadena & "Consulte su Soporte Tecnico, Soluciones Informaticas"
             MsgBox Prompt:=cadena, Buttons:=vbExclamation, Title:="Error de Registro 380:Zeus Facturacion"
             cadena = ""
        Case 3163
             MsgBox "Desvordamiento de Datos", vbCritical, "Error de Registro 3163:Zeus Facturacion"
        Case 3021
             MsgBox "No Existe Registro Activo", vbCritical, "Error de Registro 3021:Zeus Facturacion"
        Case 3200
             MsgBox "No se Puede Eliminar Tiene Datos Relacionados", vbCritical, "Error de Registro 3200:Zeus Facturacion"
        Case 3315
            MsgBox "Debe Existir una Clave Primaria", vbCritical, "Error de Registro 3315:Zeus Facturacion"
        Case 3421
            MsgBox "No se Puede Agregar este Registro.", vbInformation, "Error de Registro 3421:Zeus Facturacion"
           
        Case 3201
            MsgBox "No se Puede Modificar el Registro Desde Aqui", vbCritical, "Error de Registro 3201:Zeus Facturacion"
        Case 440
            MsgBox "No coincide la data, con la intruccion", vbCritical, "Error de Datos 440:Zeus Facturacion"
        Case 68
            MsgBox Prompt:="La unidad no está preparada. Inserte un disco en la unidad.", Buttons:=vbExclamation, Title:="Sistema Contable"
            ' Restablece la ruta a la unidad anterior.
            Drive1.Drive = Dir1.Path
            Exit Sub
        Case 20525
              MsgBox Prompt:="No se puede Leer el Reporte", Buttons:=vbExclamation, Title:="Sistema Contable"
        Case 52
             MsgBox Prompt:="La unidad no está preparada. Inserte un disco en la unidad.", Buttons:=vbExclamation, Title:="Sistema Contable"
        Case 70
             MsgBox Prompt:="La Unidad Esta protegida.Contra Escritura", Buttons:=vbExclamation, Title:="Sistema Contable"
        Case 71
             MsgBox Prompt:="La unidad no está preparada. Inserte un disco en la unidad.", Buttons:=vbExclamation, Title:="Sistema Contable"
        Case 76
            MsgBox Prompt:="No se Ha encontrado la Ruta Indicada.", Buttons:=vbExclamation, Title:="Sistema Contable"
        Case Else
            cadena = cadena & "Un error Desconocido ocurrió en este momento, Por Favor" & vbCrLf
            cadena = cadena & "si el problema persiste debe llamar al distribuidor del sistema" & vbCrLf
            cadena = cadena & "Error en la aplicación.Consulte al Soporte Tecnico Juan G. Bermúdez Tef:8502372" & vbCrLf
            MsgBox Prompt:=cadena, Buttons:=vbExclamation, Title:="Sistema Contable"
    End Select
cadena = ""
End Sub

Public Sub ColorForm(ByVal frmForm As Form)
    'Maneja las properties del form
    On Error Resume Next
    frmForm.BackColor = &H80000001
    frmForm.KeyPreview = True
    'Le da color a los frames y labels
    Dim ControlAct As Control
    For Each ControlAct In frmForm.Controls
        If TypeOf ControlAct Is Frame Then
            ControlAct.BackColor = BkColor
            ControlAct.ForeColor = FrColor
        ElseIf TypeOf ControlAct Is Label Then
            ControlAct.BackColor = BkColor
            ControlAct.ForeColor = FrColor
        ElseIf TypeOf ControlAct Is OptionButton Then
            ControlAct.BackColor = BkColor
            ControlAct.ForeColor = FrColor
        ElseIf TypeOf ControlAct Is DTPicker Then
            ControlAct.CalendarTitleBackColor = &H8000000D
            ControlAct.CalendarTitleForeColor = &H40C0&
        ElseIf TypeOf ControlAct Is SkinLabel Then
            ControlAct.BackColor = &H8000000D
            ControlAct.ForeColor = &H40C0&
        End If
        DoEvents
    Next
End Sub


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



'Function VBGetSaveFileName(FileName As String, _
'                           Optional FileTitle As String, _
'                           Optional OverWritePrompt As Boolean = True, _
'                           Optional filter As String = "All (*.*)| *.*", _
'                           Optional FilterIndex As Long = 1, _
'                           Optional InitDir As String, _
'                           Optional DlgTitle As String, _
'                           Optional DefaultExt As String, _
'                           Optional Owner As Long = -1, _
'                           Optional Flags As Long) As Boolean
'
'    Dim opfile As OpenFileName, s As String
'With opfile
'    .lStructSize = Len(opfile)
'
'    ' Add in specific flags and strip out non-VB flags
'    .Flags = (-OverWritePrompt * OFN_OVERWRITEPROMPT) Or _
'             OFN_HIDEREADONLY Or _
'             (Flags And CLng(Not (OFN_ENABLEHOOK Or _
'                                  OFN_ENABLETEMPLATE)))
'    ' Owner can take handle of owning window
'    If Owner <> -1 Then .hwndOwner = Owner
'    ' InitDir can take initial directory string
'    .lpstrInitialDir = InitDir
'    ' DefaultExt can take default extension
'    .lpstrDefExt = DefaultExt
'    ' DlgTitle can take dialog box title
'    .lpstrTitle = DlgTitle
'
'    ' Make new filter with bars (|) replacing nulls and double null at end
'    Dim ch As String, i As Integer
'    For i = 1 To Len(filter)
'        ch = Mid$(filter, i, 1)
'        If ch = "|" Or ch = ":" Then
'            s = s & vbNullChar
'        Else
'            s = s & ch
'        End If
'    Next
'    ' Put double null at end
'    s = s & vbNullChar & vbNullChar
'    .lpstrFilter = s
'    .nFilterIndex = FilterIndex
'
'    ' Pad file and file title buffers to maximum path
'    s = FileName & String$(cMaxPath - Len(FileName), 0)
'    .lpstrFile = s
'    .nMaxFile = cMaxPath
'    s = FileTitle & String$(cMaxFile - Len(FileTitle), 0)
'    .lpstrFileTitle = s
'    .nMaxFileTitle = cMaxFile
'    ' All other fields zero
'
'    If GetSaveFileName(opfile) Then
'        VBGetSaveFileName = True
'        FileName = Left$(.lpstrFile, Len(.lpstrFile))
'        FileTitle = Left$(.lpstrFileTitle, Len(.lpstrFileTitle))
'        Flags = .Flags
'        ' Return the filter index
'        FilterIndex = .nFilterIndex
'        ' Look up the filter the user selected and return that
'        filter = FilterLookup(.lpstrFilter, FilterIndex)
'    Else
'        VBGetSaveFileName = False
'        FileName = sEmpty
'        FileTitle = sEmpty
'        Flags = 0
'        FilterIndex = 0
'        filter = sEmpty
'    End If
'End With
'End Function

Function GrabaEncabezado(NumeroPeriodo As Double, NumeroTransaccion As Double, FechaTransaccion As Date, DescripcionMovimiento As String, Fuente As String, TipoMoneda As String) As Boolean


'//////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////GRABO INDICE TRANSACCIONES/////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////
 MDIPrimero.AdoConsulta.RecordSource = "SELECT  * From IndiceTransaccion WHERE (FechaTransaccion = CONVERT(DATETIME, '" & Format(FechaTransaccion, "yyyy-mm-dd") & "', 102)) AND (NumeroMovimiento = " & NumeroTransaccion & ")"
 MDIPrimero.AdoConsulta.Refresh
 If MDIPrimero.AdoConsulta.Recordset.EOF Then

                              MDIPrimero.AdoConsulta.Recordset.AddNew
                              MDIPrimero.AdoConsulta.Recordset("FechaTransaccion") = FechaTransaccion
                              If DescripcionMovimiento <> "" Then
                                MDIPrimero.AdoConsulta.Recordset("DescripcionMovimiento") = DescripcionMovimiento
                              End If
                              MDIPrimero.AdoConsulta.Recordset("NumeroMovimiento") = NumeroTransaccion
                              MDIPrimero.AdoConsulta.Recordset("Fuente") = Fuente
                              MDIPrimero.AdoConsulta.Recordset("NPeriodo") = NumeroPeriodo
                              MDIPrimero.AdoConsulta.Recordset("TipoMoneda") = TipoMoneda
                              MDIPrimero.AdoConsulta.Recordset.Update

Else
                              
                              MDIPrimero.AdoConsulta.Recordset("FechaTransaccion") = FechaTransaccion
                              If DescripcionMovimiento <> "" Then
                                MDIPrimero.AdoConsulta.Recordset("DescripcionMovimiento") = DescripcionMovimiento
                              End If
                              MDIPrimero.AdoConsulta.Recordset("NumeroMovimiento") = NumeroTransaccion
                              MDIPrimero.AdoConsulta.Recordset("Fuente") = Fuente
                              MDIPrimero.AdoConsulta.Recordset("NPeriodo") = NumeroPeriodo
                              MDIPrimero.AdoConsulta.Recordset("TipoMoneda") = TipoMoneda
                              MDIPrimero.AdoConsulta.Recordset.Update
End If

GrabaEncabezado = True

End Function

Function GrabaDetalleFactura(CodCuentas As String, FechaTransaccion As Date, NumeroTransaccion As Double, NumeroPeriodo As Double, NombreCuenta As String, DescripcionMovimiento As String, Clave As String, TasaCambio As Double, Debito As Double, Credito As Double, Fuente As String, NumeroFactura As String, FechaDescuento As Date, Descuento As Double, FechaVence As Date, CodCuentaProveedor As String, TipoFactura As String) As Boolean
   Dim DebitoAnterior As Double, CreditoAnterior As Double, ClaveAnterior As String
   Dim TipoCuenta As String
   
   TipoCuenta = ""
   MDIPrimero.AdoConsulta.RecordSource = "SELECT * From Cuentas WHERE (CodCuentas = '" & CodCuentas & "')"
   MDIPrimero.AdoConsulta.Refresh
   If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
    NombreCuenta = MDIPrimero.AdoConsulta.Recordset("DescripcionCuentas")
    TipoCuenta = MDIPrimero.AdoConsulta.Recordset("TipoCuenta")
   End If
   
   
   If Debito = 0 And Credito = 0 Then
     Exit Function
   End If
                              
                              
                              
                              '///////////////////////////////////////////////////////////////////////////////////////////
                              '/////////////////////AGREGO EL DETALLE TRANSACCION ORIGEN////////////////////////////////
                              '////////////////////////////////////////////////////////////////////////////////////////////
                              MDIPrimero.AdoConsulta.RecordSource = "SELECT * From Transacciones WHERE (CodCuentas = '" & CodCuentas & "') AND (FechaTransaccion = CONVERT(DATETIME, '" & Format(FechaTransaccion, "yyyy-mm-dd") & "', 102)) AND (NumeroMovimiento = " & NumeroTransaccion & ")"
                              MDIPrimero.AdoConsulta.Refresh

'                              If MDIPrimero.AdoConsulta.Recordset.EOF Then
                               If QUIEN <> "IVA" Then
                                MDIPrimero.AdoConsulta.Recordset.AddNew
                                 MDIPrimero.AdoConsulta.Recordset("CodCuentas") = CodCuentas
                                 MDIPrimero.AdoConsulta.Recordset("FechaTransaccion") = Format(FechaTransaccion, "dd/mm/yyyy")
                                 MDIPrimero.AdoConsulta.Recordset("NPeriodo") = NumeroPeriodo
                                 MDIPrimero.AdoConsulta.Recordset("NumeroMovimiento") = NumeroTransaccion
                                 MDIPrimero.AdoConsulta.Recordset("NombreCuenta") = NombreCuenta
                                 MDIPrimero.AdoConsulta.Recordset("DescripcionMovimiento") = DescripcionMovimiento
                                 MDIPrimero.AdoConsulta.Recordset("Clave") = Clave
                                 MDIPrimero.AdoConsulta.Recordset("TCambio") = TasaCambio
                                 MDIPrimero.AdoConsulta.Recordset("Debito") = Debito
                                 MDIPrimero.AdoConsulta.Recordset("Credito") = Credito
                                 MDIPrimero.AdoConsulta.Recordset("Fuente") = Fuente
                                 MDIPrimero.AdoConsulta.Recordset("FechaTasas") = Format(FechaTransaccion, "dd/mm/yyyy")
                                 MDIPrimero.AdoConsulta.Recordset("Beneficiario") = Beneficiario

                 
                                 If TipoFactura = "ChequePago" Then
                                   If TipoCuenta = "Bancos" Then
                                    MDIPrimero.AdoConsulta.Recordset("ChequeNo") = "#######"
                                    MDIPrimero.AdoConsulta.Recordset("FacturaNo") = NumeroFactura
                                   Else
                                    MDIPrimero.AdoConsulta.Recordset("ChequeNo") = ""
                                    MDIPrimero.AdoConsulta.Recordset("FacturaNo") = NumeroFactura
                                   End If
                                 ElseIf TipoFactura = "Recibo" Then
                                      MDIPrimero.AdoConsulta.Recordset("ChequeNo") = NumeroFactura
                                 ElseIf TipoFactura = "ReciboPago" Then
                                      MDIPrimero.AdoConsulta.Recordset("ChequeNo") = ""
                                 Else
                                      MDIPrimero.AdoConsulta.Recordset("FacturaNo") = NumeroFactura
                                   
                                 End If
                                 
                                
                                 If FechaDescuento <> "12:00:00 a.m." Then
                                  MDIPrimero.AdoConsulta.Recordset("FechaDescuento") = FechaDescuento
                                 End If
                                 MDIPrimero.AdoConsulta.Recordset("DescuentoDisponible") = Descuento
                                 If FechaVence <> "12:00:00 a.m." Then
                                  MDIPrimero.AdoConsulta.Recordset("FechaVence") = FechaVence
                                 End If
                                 MDIPrimero.AdoConsulta.Recordset("CodCuentaProveedor") = CodCuentaProveedor
                                 MDIPrimero.AdoConsulta.Recordset("TipoFactura") = TipoFactura
                                MDIPrimero.AdoConsulta.Recordset.Update
                              Else
                              
                              '///////////////////////////////////////////////////////////////////////////////////////////
                              '/////////////////////PARA EL IVA VALIDA SI LA TRANSACCION PARA EL IVA Y LA FACTURA EXSITEN////////////////////////////////
                              '////////////////////////////////////////////////////////////////////////////////////////////
                              MDIPrimero.AdoConsulta.RecordSource = "SELECT * From Transacciones WHERE (CodCuentas = '" & CodCuentas & "') AND (FechaTransaccion = CONVERT(DATETIME, '" & Format(FechaTransaccion, "yyyy-mm-dd") & "', 102)) AND (NumeroMovimiento = " & NumeroTransaccion & ") AND (FacturaNo = '" & NumeroFactura & "')"
                              MDIPrimero.AdoConsulta.Refresh

                                      If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
        
                                         DebitoAnterior = MDIPrimero.AdoConsulta.Recordset("Debito")
                                         CreditoAnterior = MDIPrimero.AdoConsulta.Recordset("Credito")
                                         ClaveAnterior = MDIPrimero.AdoConsulta.Recordset("Clave")
                                         If Clave = ClaveAnterior Then
                                           MDIPrimero.AdoConsulta.Recordset("Debito") = Debito + DebitoAnterior
                                           MDIPrimero.AdoConsulta.Recordset("Credito") = Credito + CreditoAnterior
                                           MDIPrimero.AdoConsulta.Recordset.Update
                                         Else
                                                MDIPrimero.AdoConsulta.Recordset.AddNew
                                                 MDIPrimero.AdoConsulta.Recordset("CodCuentas") = CodCuentas
                                                 MDIPrimero.AdoConsulta.Recordset("FechaTransaccion") = Format(FechaTransaccion, "dd/mm/yyyy")
                                                 MDIPrimero.AdoConsulta.Recordset("NPeriodo") = NumeroPeriodo
                                                 MDIPrimero.AdoConsulta.Recordset("NumeroMovimiento") = NumeroTransaccion
                                                 MDIPrimero.AdoConsulta.Recordset("NombreCuenta") = NombreCuenta
                                                 MDIPrimero.AdoConsulta.Recordset("DescripcionMovimiento") = DescripcionMovimiento
                                                 MDIPrimero.AdoConsulta.Recordset("Clave") = Clave
                                                 MDIPrimero.AdoConsulta.Recordset("TCambio") = TasaCambio
                                                 MDIPrimero.AdoConsulta.Recordset("Debito") = Debito
                                                 MDIPrimero.AdoConsulta.Recordset("Credito") = Credito
                                                 MDIPrimero.AdoConsulta.Recordset("Fuente") = Fuente
                                                 MDIPrimero.AdoConsulta.Recordset("FechaTasas") = Format(FechaTransaccion, "dd/mm/yyyy")
                                                 MDIPrimero.AdoConsulta.Recordset("FacturaNo") = NumeroFactura
                                                 If FechaDescuento <> "12:00:00 a.m." Then
                                                  MDIPrimero.AdoConsulta.Recordset("FechaDescuento") = FechaDescuento
                                                 End If
                                                 MDIPrimero.AdoConsulta.Recordset("DescuentoDisponible") = Descuento
                                                 If FechaVence <> "12:00:00 a.m." Then
                                                  MDIPrimero.AdoConsulta.Recordset("FechaVence") = FechaVence
                                                 End If
                                                 MDIPrimero.AdoConsulta.Recordset("CodCuentaProveedor") = CodCuentaProveedor
                                                 MDIPrimero.AdoConsulta.Recordset("TipoFactura") = TipoFactura
                                                MDIPrimero.AdoConsulta.Recordset.Update
                                         End If
                                      Else
                                                MDIPrimero.AdoConsulta.Recordset.AddNew
                                                 MDIPrimero.AdoConsulta.Recordset("CodCuentas") = CodCuentas
                                                 MDIPrimero.AdoConsulta.Recordset("FechaTransaccion") = Format(FechaTransaccion, "dd/mm/yyyy")
                                                 MDIPrimero.AdoConsulta.Recordset("NPeriodo") = NumeroPeriodo
                                                 MDIPrimero.AdoConsulta.Recordset("NumeroMovimiento") = NumeroTransaccion
                                                 MDIPrimero.AdoConsulta.Recordset("NombreCuenta") = NombreCuenta
                                                 MDIPrimero.AdoConsulta.Recordset("DescripcionMovimiento") = DescripcionMovimiento
                                                 MDIPrimero.AdoConsulta.Recordset("Clave") = Clave
                                                 MDIPrimero.AdoConsulta.Recordset("TCambio") = TasaCambio
                                                 MDIPrimero.AdoConsulta.Recordset("Debito") = Debito
                                                 MDIPrimero.AdoConsulta.Recordset("Credito") = Credito
                                                 MDIPrimero.AdoConsulta.Recordset("Fuente") = Fuente
                                                 MDIPrimero.AdoConsulta.Recordset("FechaTasas") = Format(FechaTransaccion, "dd/mm/yyyy")
                                                 MDIPrimero.AdoConsulta.Recordset("FacturaNo") = NumeroFactura
                                                 If FechaDescuento <> "12:00:00 a.m." Then
                                                  MDIPrimero.AdoConsulta.Recordset("FechaDescuento") = FechaDescuento
                                                 End If
                                                 MDIPrimero.AdoConsulta.Recordset("DescuentoDisponible") = Descuento
                                                 If FechaVence <> "12:00:00 a.m." Then
                                                  MDIPrimero.AdoConsulta.Recordset("FechaVence") = FechaVence
                                                 End If
                                                 MDIPrimero.AdoConsulta.Recordset("CodCuentaProveedor") = CodCuentaProveedor
                                                 MDIPrimero.AdoConsulta.Recordset("TipoFactura") = TipoFactura
                                                MDIPrimero.AdoConsulta.Recordset.Update
                                      End If
                                  End If
                                

End Function
Function GrabaDetalleNomina(CodCuentas As String, FechaTransaccion As Date, NumeroTransaccion As Double, NumeroPeriodo As Double, NombreCuenta As String, DescripcionMovimiento As String, Clave As String, TasaCambio As Double, Debito As Double, Credito As Double, Fuente As String, NumeroFactura As String, FechaDescuento As Date, Descuento As Double, FechaVence As Date, CodCuentaProveedor As String, TipoFactura As String) As Boolean
   Dim TipoCuenta As String, NombreEmpleado As String, TipoMoneda As String
   
   MDIPrimero.AdoConsulta.RecordSource = "SELECT * From Cuentas WHERE (CodCuentas = '" & CodCuentas & "')"
   MDIPrimero.AdoConsulta.Refresh
   If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
    NombreCuenta = MDIPrimero.AdoConsulta.Recordset("DescripcionCuentas")
    TipoCuenta = MDIPrimero.AdoConsulta.Recordset("TipoCuenta")
   End If
   
   
   If Debito = 0 And Credito = 0 Then
     Exit Function
   End If
   
   
   If Fuente = "CHEQUE" Then
            
'            NumeroFactura = "-"
            TipoMoneda = "Córdobas"

   
   
            '///////////si el cheque no se ha grabado, guardo el numero Voucher/////////////////

              If TipoCuenta = "Bancos" Then
            
                    MDIPrimero.AdoConsulta.RecordSource = "SELECT NConsecutivoVoucher.CodCuenta, NConsecutivoVoucher.ConsecutivoVoucher, NConsecutivoVoucher.NPeriodo From NConsecutivoVoucher Where (((NConsecutivoVoucher.CodCuenta) = '" & CodigoCuenta & "') And ((NConsecutivoVoucher.NPeriodo) = " & NumeroPeriodo & "))"
                    MDIPrimero.AdoConsulta.Refresh
                    If MDIPrimero.AdoConsulta.Recordset.EOF Then
                       MDIPrimero.AdoConsulta.Recordset.AddNew
                         MDIPrimero.AdoConsulta.Recordset("CodCuenta") = CodigoCuenta
                         MDIPrimero.AdoConsulta.Recordset("NPeriodo") = NumeroPeriodo
                         MDIPrimero.AdoConsulta.Recordset("ConsecutivoVoucher") = 1
                       MDIPrimero.AdoConsulta.Recordset.Update
                       NumeroVoucher = 1
                    Else
                       'MDIPrimero.'AdoConsulta.Recordset.Edit
                        MDIPrimero.AdoConsulta.Recordset("ConsecutivoVoucher") = MDIPrimero.AdoConsulta.Recordset("ConsecutivoVoucher") + 1
                       MDIPrimero.AdoConsulta.Recordset.Update
                     NumeroVoucher = MDIPrimero.AdoConsulta.Recordset("ConsecutivoVoucher")
                    End If

                    ConsecutivoVoucher = Month(FechaTransaccion)
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

             
                              '///////////////////////////////////////////////////////////////////////////////////////////
                              '/////////////////////AGREGO EL DETALLE TRANSACCION ORIGEN////////////////////////////////
                              '////////////////////////////////////////////////////////////////////////////////////////////
                              MDIPrimero.AdoConsulta.RecordSource = "SELECT * From Transacciones WHERE (CodCuentas = '" & CodCuentas & "') AND (FechaTransaccion = CONVERT(DATETIME, '" & Format(FechaTransaccion, "yyyy-mm-dd") & "', 102)) AND (NumeroMovimiento = " & NumeroTransaccion & ")  AND (Clave = '" & Clave & "')"
                              MDIPrimero.AdoConsulta.Refresh

                              If MDIPrimero.AdoConsulta.Recordset.EOF Then
                                MDIPrimero.AdoConsulta.Recordset.AddNew
                                 MDIPrimero.AdoConsulta.Recordset("CodCuentas") = CodCuentas
                                 MDIPrimero.AdoConsulta.Recordset("FechaTransaccion") = Format(FechaTransaccion, "dd/mm/yyyy")
                                 MDIPrimero.AdoConsulta.Recordset("NPeriodo") = NumeroPeriodo
                                 MDIPrimero.AdoConsulta.Recordset("NumeroMovimiento") = NumeroTransaccion
                                 MDIPrimero.AdoConsulta.Recordset("NombreCuenta") = NombreCuenta
                                 
                                 If TipoCuenta = "Bancos" Then
                                  MDIPrimero.AdoConsulta.Recordset("ChequeNo") = NumeroFactura
                                  MDIPrimero.AdoConsulta.Recordset("DescripcionMovimiento") = DescripcionMovimiento & " " & TipoFactura
                                 Else
                                  MDIPrimero.AdoConsulta.Recordset("ChequeNo") = "-"
                                  MDIPrimero.AdoConsulta.Recordset("DescripcionMovimiento") = DescripcionMovimiento
                                 End If
                                 MDIPrimero.AdoConsulta.Recordset("Clave") = Clave
                                 MDIPrimero.AdoConsulta.Recordset("TCambio") = TasaCambio
                                 MDIPrimero.AdoConsulta.Recordset("Debito") = Format(Debito, "####0.00")
                                 MDIPrimero.AdoConsulta.Recordset("Credito") = Format(Credito, "####0.00")
                                 MDIPrimero.AdoConsulta.Recordset("Fuente") = Fuente
                                 MDIPrimero.AdoConsulta.Recordset("VoucherNo") = cadena
                                 MDIPrimero.AdoConsulta.Recordset("FechaTasas") = Format(FechaTransaccion, "dd/mm/yyyy")
                                 MDIPrimero.AdoConsulta.Recordset("Beneficiario") = TipoFactura
                                MDIPrimero.AdoConsulta.Recordset.Update
                              End If
                              
    Else
                              
                              
                              
                              '///////////////////////////////////////////////////////////////////////////////////////////
                              '/////////////////////AGREGO EL DETALLE TRANSACCION ORIGEN////////////////////////////////
                              '////////////////////////////////////////////////////////////////////////////////////////////
                              MDIPrimero.AdoConsulta.RecordSource = "SELECT * From Transacciones WHERE (CodCuentas = '" & CodCuentas & "') AND (FechaTransaccion = CONVERT(DATETIME, '" & Format(FechaTransaccion, "yyyy-mm-dd") & "', 102)) AND (NumeroMovimiento = " & NumeroTransaccion & ")  AND (Clave = '" & Clave & "')"
                              MDIPrimero.AdoConsulta.Refresh

                              If MDIPrimero.AdoConsulta.Recordset.EOF Then
                                MDIPrimero.AdoConsulta.Recordset.AddNew
                                 MDIPrimero.AdoConsulta.Recordset("CodCuentas") = CodCuentas
                                 MDIPrimero.AdoConsulta.Recordset("FechaTransaccion") = Format(FechaTransaccion, "dd/mm/yyyy")
                                 MDIPrimero.AdoConsulta.Recordset("NPeriodo") = NumeroPeriodo
                                 MDIPrimero.AdoConsulta.Recordset("NumeroMovimiento") = NumeroTransaccion
                                 MDIPrimero.AdoConsulta.Recordset("NombreCuenta") = NombreCuenta
                                 MDIPrimero.AdoConsulta.Recordset("DescripcionMovimiento") = DescripcionMovimiento
                                 MDIPrimero.AdoConsulta.Recordset("Clave") = Clave
                                 MDIPrimero.AdoConsulta.Recordset("TCambio") = TasaCambio
                                 MDIPrimero.AdoConsulta.Recordset("Debito") = Debito
                                 MDIPrimero.AdoConsulta.Recordset("Credito") = Credito
                                 MDIPrimero.AdoConsulta.Recordset("Fuente") = Fuente
                                 MDIPrimero.AdoConsulta.Recordset("FechaTasas") = Format(FechaTransaccion, "dd/mm/yyyy")
                                 If NumeroFactura = "#######" Then
                                   MDIPrimero.AdoConsulta.Recordset("ChequeNo") = NumeroFactura
                                   NumeroFactura = "-"
                                 End If
                                 MDIPrimero.AdoConsulta.Recordset("FacturaNo") = NumeroFactura
                                 If FechaDescuento <> "12:00:00 a.m." Then
                                  MDIPrimero.AdoConsulta.Recordset("FechaDescuento") = FechaDescuento
                                 End If
                                 MDIPrimero.AdoConsulta.Recordset("DescuentoDisponible") = Descuento
                                 If FechaVence <> "12:00:00 a.m." Then
                                  MDIPrimero.AdoConsulta.Recordset("FechaVence") = FechaVence
                                 End If
                                 MDIPrimero.AdoConsulta.Recordset("CodCuentaProveedor") = CodCuentaProveedor
                                 MDIPrimero.AdoConsulta.Recordset("TipoFactura") = TipoFactura
                                MDIPrimero.AdoConsulta.Recordset.Update
                              Else
'                                 MDIPrimero.AdoConsulta.Recordset("DescripcionMovimiento") = DescripcionMovimiento
'                                 MDIPrimero.AdoConsulta.Recordset("Clave") = Clave
'                                 MDIPrimero.AdoConsulta.Recordset("TCambio") = TasaCambio
'                                 MDIPrimero.AdoConsulta.Recordset("Debito") = Debito
'                                 MDIPrimero.AdoConsulta.Recordset("Credito") = Credito
'                                 MDIPrimero.AdoConsulta.Recordset("Fuente") = Fuente
'                                 MDIPrimero.AdoConsulta.Recordset("FechaTasas") = FechaTransaccion
'                                 MDIPrimero.AdoConsulta.Recordset("FacturaNo") = NumeroFactura
'                                 MDIPrimero.AdoConsulta.Recordset("FechaDescuento") = FechaDescuento
'                                 MDIPrimero.AdoConsulta.Recordset("DescuentoDisponible") = Descuento
'                                 MDIPrimero.AdoConsulta.Recordset("FechaVence") = FechaVence
'                                 MDIPrimero.AdoConsulta.Recordset("CodCuentaProveedor") = CodCuentaProveedor
'                                 MDIPrimero.AdoConsulta.Recordset("TipoFactura") = TipoFactura
                                 MDIPrimero.AdoConsulta.Recordset("Debito") = Format(Debito, "##,##0.00") + MDIPrimero.AdoConsulta.Recordset("Debito")
                                 MDIPrimero.AdoConsulta.Recordset("Credito") = Format(Credito, "##,##0.00") + MDIPrimero.AdoConsulta.Recordset("Credito")
                                 MDIPrimero.AdoConsulta.Recordset.Update
                              End If
                              
                              
     End If
                                

End Function
Function GrabaDetalleCheque(CodCuentas As String, FechaTransaccion As Date, NumeroTransaccion As Double, NumeroPeriodo As Double, NombreCuenta As String, DescripcionMovimiento As String, Clave As String, TasaCambio As Double, Debito As Double, Credito As Double, Fuente As String, NumeroFactura As String, FechaDescuento As Date, Descuento As Double, FechaVence As Date, CodCuentaProveedor As String, TipoFactura As String, VoucherNo As String) As Boolean
   Dim TipoCuenta As String, NombreEmpleado As String, TipoMoneda As String
   
   MDIPrimero.AdoConsulta.RecordSource = "SELECT * From Cuentas WHERE (CodCuentas = '" & CodCuentas & "')"
   MDIPrimero.AdoConsulta.Refresh
   If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
    NombreCuenta = MDIPrimero.AdoConsulta.Recordset("DescripcionCuentas")
    TipoCuenta = MDIPrimero.AdoConsulta.Recordset("TipoCuenta")
   End If
   
   
   If Debito = 0 And Credito = 0 Then
     Exit Function
   End If
   
   
   If Fuente = "CHEQUE" Then
            
'            NumeroFactura = "-"
'            TipoMoneda = "Córdobas"

   
   
            '///////////si el cheque no se ha grabado, guardo el numero Voucher/////////////////

              cadena = VoucherNo
   
              If TipoCuenta = "Bancos" Then
            
                    MDIPrimero.AdoConsulta.RecordSource = "SELECT NConsecutivoVoucher.CodCuenta, NConsecutivoVoucher.ConsecutivoVoucher, NConsecutivoVoucher.NPeriodo From NConsecutivoVoucher Where (((NConsecutivoVoucher.CodCuenta) = '" & CodigoCuenta & "') And ((NConsecutivoVoucher.NPeriodo) = " & NumeroPeriodo & "))"
                    MDIPrimero.AdoConsulta.Refresh
                    If MDIPrimero.AdoConsulta.Recordset.EOF Then
                       MDIPrimero.AdoConsulta.Recordset.AddNew
                         MDIPrimero.AdoConsulta.Recordset("CodCuenta") = CodigoCuenta
                         MDIPrimero.AdoConsulta.Recordset("NPeriodo") = NumeroPeriodo
                         MDIPrimero.AdoConsulta.Recordset("ConsecutivoVoucher") = 1
                       MDIPrimero.AdoConsulta.Recordset.Update
                       NumeroVoucher = 1
                    Else
                       'MDIPrimero.'AdoConsulta.Recordset.Edit
                        MDIPrimero.AdoConsulta.Recordset("ConsecutivoVoucher") = MDIPrimero.AdoConsulta.Recordset("ConsecutivoVoucher") + 1
                       MDIPrimero.AdoConsulta.Recordset.Update
                     NumeroVoucher = MDIPrimero.AdoConsulta.Recordset("ConsecutivoVoucher")
                    End If

                    ConsecutivoVoucher = Month(FechaTransaccion)
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

             
                              '///////////////////////////////////////////////////////////////////////////////////////////
                              '/////////////////////AGREGO EL DETALLE TRANSACCION ORIGEN////////////////////////////////
                              '////////////////////////////////////////////////////////////////////////////////////////////
                              MDIPrimero.AdoConsulta.RecordSource = "SELECT * From Transacciones WHERE (CodCuentas = '" & CodCuentas & "') AND (FechaTransaccion = CONVERT(DATETIME, '" & Format(FechaTransaccion, "yyyy-mm-dd") & "', 102)) AND (NumeroMovimiento = " & NumeroTransaccion & ")  AND (Clave = '" & Clave & "')"
                              MDIPrimero.AdoConsulta.Refresh

                              If MDIPrimero.AdoConsulta.Recordset.EOF Then
                                MDIPrimero.AdoConsulta.Recordset.AddNew
                                 MDIPrimero.AdoConsulta.Recordset("CodCuentas") = CodCuentas
                                 MDIPrimero.AdoConsulta.Recordset("FechaTransaccion") = Format(FechaTransaccion, "dd/mm/yyyy")
                                 MDIPrimero.AdoConsulta.Recordset("NPeriodo") = NumeroPeriodo
                                 MDIPrimero.AdoConsulta.Recordset("NumeroMovimiento") = NumeroTransaccion
                                 MDIPrimero.AdoConsulta.Recordset("NombreCuenta") = NombreCuenta
                                 
                                 If TipoCuenta = "Bancos" Then
'                                  If NumeroFactura = "-" Then
'                                    MDIPrimero.AdoConsulta.Recordset("ChequeNo") = "#######"
'                                  Else
'                                    MDIPrimero.AdoConsulta.Recordset("ChequeNo") = NumeroFactura
'                                  End If
                                  MDIPrimero.AdoConsulta.Recordset("ChequeNo") = "#######"
                                  MDIPrimero.AdoConsulta.Recordset("DescripcionMovimiento") = DescripcionMovimiento & " " & TipoFactura
                                 Else
                                  MDIPrimero.AdoConsulta.Recordset("FacturaNo") = NumeroFactura
                                  MDIPrimero.AdoConsulta.Recordset("ChequeNo") = "-"
                                  MDIPrimero.AdoConsulta.Recordset("DescripcionMovimiento") = DescripcionMovimiento
                                 End If
                                 MDIPrimero.AdoConsulta.Recordset("Clave") = Clave
                                 MDIPrimero.AdoConsulta.Recordset("TCambio") = TasaCambio
                                 MDIPrimero.AdoConsulta.Recordset("Debito") = Format(Debito, "####0.00")
                                 MDIPrimero.AdoConsulta.Recordset("Credito") = Format(Credito, "####0.00")
                                 MDIPrimero.AdoConsulta.Recordset("Fuente") = Fuente
                                 MDIPrimero.AdoConsulta.Recordset("VoucherNo") = cadena
                                 MDIPrimero.AdoConsulta.Recordset("FechaTasas") = Format(FechaTransaccion, "dd/mm/yyyy")
                                 MDIPrimero.AdoConsulta.Recordset("Beneficiario") = TipoFactura
                                MDIPrimero.AdoConsulta.Recordset.Update
                              End If
                              
    Else
                              
                              
                              
                              '///////////////////////////////////////////////////////////////////////////////////////////
                              '/////////////////////AGREGO EL DETALLE TRANSACCION ORIGEN////////////////////////////////
                              '////////////////////////////////////////////////////////////////////////////////////////////
                              MDIPrimero.AdoConsulta.RecordSource = "SELECT * From Transacciones WHERE (CodCuentas = '" & CodCuentas & "') AND (FechaTransaccion = CONVERT(DATETIME, '" & Format(FechaTransaccion, "yyyy-mm-dd") & "', 102)) AND (NumeroMovimiento = " & NumeroTransaccion & ")  AND (Clave = '" & Clave & "')"
                              MDIPrimero.AdoConsulta.Refresh

                              If MDIPrimero.AdoConsulta.Recordset.EOF Then
                                MDIPrimero.AdoConsulta.Recordset.AddNew
                                 MDIPrimero.AdoConsulta.Recordset("CodCuentas") = CodCuentas
                                 MDIPrimero.AdoConsulta.Recordset("FechaTransaccion") = Format(FechaTransaccion, "dd/mm/yyyy")
                                 MDIPrimero.AdoConsulta.Recordset("NPeriodo") = NumeroPeriodo
                                 MDIPrimero.AdoConsulta.Recordset("NumeroMovimiento") = NumeroTransaccion
                                 MDIPrimero.AdoConsulta.Recordset("NombreCuenta") = NombreCuenta
                                 MDIPrimero.AdoConsulta.Recordset("DescripcionMovimiento") = DescripcionMovimiento
                                 MDIPrimero.AdoConsulta.Recordset("Clave") = Clave
                                 MDIPrimero.AdoConsulta.Recordset("TCambio") = TasaCambio
                                 MDIPrimero.AdoConsulta.Recordset("Debito") = Debito
                                 MDIPrimero.AdoConsulta.Recordset("Credito") = Credito
                                 MDIPrimero.AdoConsulta.Recordset("Fuente") = Fuente
                                 MDIPrimero.AdoConsulta.Recordset("FechaTasas") = Format(FechaTransaccion, "dd/mm/yyyy")
                                 If NumeroFactura = "#######" Then
                                   MDIPrimero.AdoConsulta.Recordset("ChequeNo") = NumeroFactura
                                   NumeroFactura = "-"
                                 End If
                                 MDIPrimero.AdoConsulta.Recordset("FacturaNo") = NumeroFactura
                                 If FechaDescuento <> "12:00:00 a.m." Then
                                  MDIPrimero.AdoConsulta.Recordset("FechaDescuento") = FechaDescuento
                                 End If
                                 MDIPrimero.AdoConsulta.Recordset("DescuentoDisponible") = Descuento
                                 If FechaVence <> "12:00:00 a.m." Then
                                  MDIPrimero.AdoConsulta.Recordset("FechaVence") = FechaVence
                                 End If
                                 MDIPrimero.AdoConsulta.Recordset("CodCuentaProveedor") = CodCuentaProveedor
                                 MDIPrimero.AdoConsulta.Recordset("TipoFactura") = TipoFactura
                                MDIPrimero.AdoConsulta.Recordset.Update
                              Else
'                                 MDIPrimero.AdoConsulta.Recordset("DescripcionMovimiento") = DescripcionMovimiento
'                                 MDIPrimero.AdoConsulta.Recordset("Clave") = Clave
'                                 MDIPrimero.AdoConsulta.Recordset("TCambio") = TasaCambio
'                                 MDIPrimero.AdoConsulta.Recordset("Debito") = Debito
'                                 MDIPrimero.AdoConsulta.Recordset("Credito") = Credito
'                                 MDIPrimero.AdoConsulta.Recordset("Fuente") = Fuente
'                                 MDIPrimero.AdoConsulta.Recordset("FechaTasas") = FechaTransaccion
'                                 MDIPrimero.AdoConsulta.Recordset("FacturaNo") = NumeroFactura
'                                 MDIPrimero.AdoConsulta.Recordset("FechaDescuento") = FechaDescuento
'                                 MDIPrimero.AdoConsulta.Recordset("DescuentoDisponible") = Descuento
'                                 MDIPrimero.AdoConsulta.Recordset("FechaVence") = FechaVence
'                                 MDIPrimero.AdoConsulta.Recordset("CodCuentaProveedor") = CodCuentaProveedor
'                                 MDIPrimero.AdoConsulta.Recordset("TipoFactura") = TipoFactura
                                 MDIPrimero.AdoConsulta.Recordset("Debito") = Format(Debito, "##,##0.00") + MDIPrimero.AdoConsulta.Recordset("Debito")
                                 MDIPrimero.AdoConsulta.Recordset("Credito") = Format(Credito, "##,##0.00") + MDIPrimero.AdoConsulta.Recordset("Credito")
                                 MDIPrimero.AdoConsulta.Recordset.Update
                              End If
                              
                              
     End If
                                

End Function


Function GrabaDetalleChequeSolicitud(CodCuentas As String, FechaTransaccion As Date, NumeroTransaccion As Double, NumeroPeriodo As Double, NombreCuenta As String, DescripcionMovimiento As String, Clave As String, TasaCambio As Double, Debito As Double, Credito As Double, Fuente As String, NumeroFactura As String, FechaDescuento As Date, Descuento As Double, FechaVence As Date, CodCuentaProveedor As String, TipoFactura As String, VoucherNo As String, KeyPresupuesto As String, Presupuesto As String) As Boolean
   Dim TipoCuenta As String, NombreEmpleado As String, TipoMoneda As String
   
   MDIPrimero.AdoConsulta.RecordSource = "SELECT * From Cuentas WHERE (CodCuentas = '" & CodCuentas & "')"
   MDIPrimero.AdoConsulta.Refresh
   If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
    NombreCuenta = MDIPrimero.AdoConsulta.Recordset("DescripcionCuentas")
    TipoCuenta = MDIPrimero.AdoConsulta.Recordset("TipoCuenta")
   End If
   
   
   If Debito = 0 And Credito = 0 Then
     Exit Function
   End If
   
   
   
   
   
   If Fuente = "CHEQUE" Then
            
'            NumeroFactura = "-"
'            TipoMoneda = "Córdobas"

   
   
            '///////////si el cheque no se ha grabado, guardo el numero Voucher/////////////////

              cadena = VoucherNo
   
              If TipoCuenta = "Bancos" Then
            
                    MDIPrimero.AdoConsulta.RecordSource = "SELECT NConsecutivoVoucher.CodCuenta, NConsecutivoVoucher.ConsecutivoVoucher, NConsecutivoVoucher.NPeriodo From NConsecutivoVoucher Where (((NConsecutivoVoucher.CodCuenta) = '" & CodigoCuenta & "') And ((NConsecutivoVoucher.NPeriodo) = " & NumeroPeriodo & "))"
                    MDIPrimero.AdoConsulta.Refresh
                    If MDIPrimero.AdoConsulta.Recordset.EOF Then
                       MDIPrimero.AdoConsulta.Recordset.AddNew
                         MDIPrimero.AdoConsulta.Recordset("CodCuenta") = CodigoCuenta
                         MDIPrimero.AdoConsulta.Recordset("NPeriodo") = NumeroPeriodo
                         MDIPrimero.AdoConsulta.Recordset("ConsecutivoVoucher") = 1
                       MDIPrimero.AdoConsulta.Recordset.Update
                       NumeroVoucher = 1
                    Else
                       'MDIPrimero.'AdoConsulta.Recordset.Edit
                        MDIPrimero.AdoConsulta.Recordset("ConsecutivoVoucher") = MDIPrimero.AdoConsulta.Recordset("ConsecutivoVoucher") + 1
                       MDIPrimero.AdoConsulta.Recordset.Update
                     NumeroVoucher = MDIPrimero.AdoConsulta.Recordset("ConsecutivoVoucher")
                    End If

                    ConsecutivoVoucher = Month(FechaTransaccion)
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

             
                              '///////////////////////////////////////////////////////////////////////////////////////////
                              '/////////////////////AGREGO EL DETALLE TRANSACCION ORIGEN////////////////////////////////
                              '////////////////////////////////////////////////////////////////////////////////////////////
                              MDIPrimero.AdoConsulta.RecordSource = "SELECT * From Transacciones WHERE (CodCuentas = '" & CodCuentas & "') AND (FechaTransaccion = CONVERT(DATETIME, '" & Format(FechaTransaccion, "yyyy-mm-dd") & "', 102)) AND (NumeroMovimiento = " & NumeroTransaccion & ")  AND (Clave = '" & Clave & "')"
                              MDIPrimero.AdoConsulta.Refresh

                              If MDIPrimero.AdoConsulta.Recordset.EOF Then
                                MDIPrimero.AdoConsulta.Recordset.AddNew
                                 MDIPrimero.AdoConsulta.Recordset("CodCuentas") = CodCuentas
                                 MDIPrimero.AdoConsulta.Recordset("FechaTransaccion") = Format(FechaTransaccion, "dd/mm/yyyy")
                                 MDIPrimero.AdoConsulta.Recordset("NPeriodo") = NumeroPeriodo
                                 MDIPrimero.AdoConsulta.Recordset("NumeroMovimiento") = NumeroTransaccion
                                 MDIPrimero.AdoConsulta.Recordset("NombreCuenta") = NombreCuenta
                                 
                                 If TipoCuenta = "Bancos" Then
'                                  If NumeroFactura = "-" Then
'                                    MDIPrimero.AdoConsulta.Recordset("ChequeNo") = "#######"
'                                  Else
'                                    MDIPrimero.AdoConsulta.Recordset("ChequeNo") = NumeroFactura
'                                  End If
                                  MDIPrimero.AdoConsulta.Recordset("ChequeNo") = "#######"
                                  MDIPrimero.AdoConsulta.Recordset("DescripcionMovimiento") = DescripcionMovimiento & " " & TipoFactura
                                 Else
                                  MDIPrimero.AdoConsulta.Recordset("FacturaNo") = NumeroFactura
                                  MDIPrimero.AdoConsulta.Recordset("ChequeNo") = "-"
                                  MDIPrimero.AdoConsulta.Recordset("DescripcionMovimiento") = DescripcionMovimiento
                                 End If
                                 MDIPrimero.AdoConsulta.Recordset("Clave") = Clave
                                 MDIPrimero.AdoConsulta.Recordset("TCambio") = TasaCambio
                                 MDIPrimero.AdoConsulta.Recordset("Debito") = Format(Debito, "####0.00")
                                 MDIPrimero.AdoConsulta.Recordset("Credito") = Format(Credito, "####0.00")
                                 MDIPrimero.AdoConsulta.Recordset("Fuente") = Fuente
                                 MDIPrimero.AdoConsulta.Recordset("VoucherNo") = cadena
                                 MDIPrimero.AdoConsulta.Recordset("FechaTasas") = Format(FechaTransaccion, "dd/mm/yyyy")
                                 MDIPrimero.AdoConsulta.Recordset("Beneficiario") = TipoFactura
                                 MDIPrimero.AdoConsulta.Recordset("KeyPresupuesto") = KeyPresupuesto
                                 MDIPrimero.AdoConsulta.Recordset("Presupuesto") = Presupuesto
                                MDIPrimero.AdoConsulta.Recordset.Update
                              End If
                              
    Else
                              
                              
                              
                              '///////////////////////////////////////////////////////////////////////////////////////////
                              '/////////////////////AGREGO EL DETALLE TRANSACCION ORIGEN////////////////////////////////
                              '////////////////////////////////////////////////////////////////////////////////////////////
                              MDIPrimero.AdoConsulta.RecordSource = "SELECT * From Transacciones WHERE (CodCuentas = '" & CodCuentas & "') AND (FechaTransaccion = CONVERT(DATETIME, '" & Format(FechaTransaccion, "yyyy-mm-dd") & "', 102)) AND (NumeroMovimiento = " & NumeroTransaccion & ")  AND (Clave = '" & Clave & "')"
                              MDIPrimero.AdoConsulta.Refresh

                              If MDIPrimero.AdoConsulta.Recordset.EOF Then
                                MDIPrimero.AdoConsulta.Recordset.AddNew
                                 MDIPrimero.AdoConsulta.Recordset("CodCuentas") = CodCuentas
                                 MDIPrimero.AdoConsulta.Recordset("FechaTransaccion") = Format(FechaTransaccion, "dd/mm/yyyy")
                                 MDIPrimero.AdoConsulta.Recordset("NPeriodo") = NumeroPeriodo
                                 MDIPrimero.AdoConsulta.Recordset("NumeroMovimiento") = NumeroTransaccion
                                 MDIPrimero.AdoConsulta.Recordset("NombreCuenta") = NombreCuenta
                                 MDIPrimero.AdoConsulta.Recordset("DescripcionMovimiento") = DescripcionMovimiento
                                 MDIPrimero.AdoConsulta.Recordset("Clave") = Clave
                                 MDIPrimero.AdoConsulta.Recordset("TCambio") = TasaCambio
                                 MDIPrimero.AdoConsulta.Recordset("Debito") = Debito
                                 MDIPrimero.AdoConsulta.Recordset("Credito") = Credito
                                 MDIPrimero.AdoConsulta.Recordset("Fuente") = Fuente
                                 MDIPrimero.AdoConsulta.Recordset("FechaTasas") = Format(FechaTransaccion, "dd/mm/yyyy")
                                 If NumeroFactura = "#######" Then
                                   MDIPrimero.AdoConsulta.Recordset("ChequeNo") = NumeroFactura
                                   NumeroFactura = "-"
                                 End If
                                 MDIPrimero.AdoConsulta.Recordset("FacturaNo") = NumeroFactura
                                 If FechaDescuento <> "12:00:00 a.m." Then
                                  MDIPrimero.AdoConsulta.Recordset("FechaDescuento") = FechaDescuento
                                 End If
                                 MDIPrimero.AdoConsulta.Recordset("DescuentoDisponible") = Descuento
                                 If FechaVence <> "12:00:00 a.m." Then
                                  MDIPrimero.AdoConsulta.Recordset("FechaVence") = FechaVence
                                 End If
                                 MDIPrimero.AdoConsulta.Recordset("CodCuentaProveedor") = CodCuentaProveedor
                                 MDIPrimero.AdoConsulta.Recordset("TipoFactura") = TipoFactura
                                 MDIPrimero.AdoConsulta.Recordset("KeyPresupuesto") = KeyPresupuesto
                                 MDIPrimero.AdoConsulta.Recordset("Presupuesto") = Presupuesto
                                MDIPrimero.AdoConsulta.Recordset.Update
                              Else
'                                 MDIPrimero.AdoConsulta.Recordset("DescripcionMovimiento") = DescripcionMovimiento
'                                 MDIPrimero.AdoConsulta.Recordset("Clave") = Clave
'                                 MDIPrimero.AdoConsulta.Recordset("TCambio") = TasaCambio
'                                 MDIPrimero.AdoConsulta.Recordset("Debito") = Debito
'                                 MDIPrimero.AdoConsulta.Recordset("Credito") = Credito
'                                 MDIPrimero.AdoConsulta.Recordset("Fuente") = Fuente
'                                 MDIPrimero.AdoConsulta.Recordset("FechaTasas") = FechaTransaccion
'                                 MDIPrimero.AdoConsulta.Recordset("FacturaNo") = NumeroFactura
'                                 MDIPrimero.AdoConsulta.Recordset("FechaDescuento") = FechaDescuento
'                                 MDIPrimero.AdoConsulta.Recordset("DescuentoDisponible") = Descuento
'                                 MDIPrimero.AdoConsulta.Recordset("FechaVence") = FechaVence
'                                 MDIPrimero.AdoConsulta.Recordset("CodCuentaProveedor") = CodCuentaProveedor
'                                 MDIPrimero.AdoConsulta.Recordset("TipoFactura") = TipoFactura
                                 MDIPrimero.AdoConsulta.Recordset("Debito") = Format(Debito, "##,##0.00") + MDIPrimero.AdoConsulta.Recordset("Debito")
                                 MDIPrimero.AdoConsulta.Recordset("Credito") = Format(Credito, "##,##0.00") + MDIPrimero.AdoConsulta.Recordset("Credito")
                                 MDIPrimero.AdoConsulta.Recordset.Update
                              End If
                              
                              
     End If
                                

End Function





Function BuscaCuenta(CodigoCuenta As String) As String
  Dim SqlString As String
  
        SqlString = "SELECT * From Cuentas WHERE (CodCuentas = '" & CodigoCuenta & "')"
        MDIPrimero.AdoConsulta.RecordSource = SqlString
        MDIPrimero.AdoConsulta.Refresh
        If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
           BuscaCuenta = MDIPrimero.AdoConsulta.Recordset("DescripcionCuentas")
        Else
           BuscaCuenta = "Nulo"
        End If

End Function

Function BuscaCodigoProducto(CodigoProducto As String) As String


    '///////////////////////////////////////////////////////////////////////////////////////////////
    '//////////////////////BUSCO LA CUENTA DEL PRODUCTO//////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////
            
            SqlString = "SELECT  * From Productos WHERE (Cod_Productos = '" & CodigoProducto & "')"
            FrmContabilizaFacturacion.AdoConsultaFactura.RecordSource = SqlString
            FrmContabilizaFacturacion.AdoConsultaFactura.Refresh
         If Not FrmContabilizaFacturacion.AdoConsultaFactura.Recordset.EOF Then
               If Not IsNull(FrmContabilizaFacturacion.AdoConsultaFactura.Recordset("Cod_Cuenta_Ventas")) Then
                      BuscaCodigoProducto = FrmContabilizaFacturacion.AdoConsultaFactura.Recordset("Cod_Cuenta_Ventas")
               End If
         End If

     
End Function
Function BuscaCuentaImpuestos(CodigoImpuesto As String) As String

        BuscaCuentaImpuestos = "Nulo"
    '///////////////////////////////////////////////////////////////////////////////////////////////
    '//////////////////////BUSCO LA CUENTA DEL PRODUCTO//////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////
            SqlString = "SELECT * From Impuestos WHERE (Cod_Iva = '" & CodigoImpuesto & "')"
            FrmContabilizaFacturacion.AdoConsultaFactura.RecordSource = SqlString
            FrmContabilizaFacturacion.AdoConsultaFactura.Refresh
         If Not FrmContabilizaFacturacion.AdoConsultaFactura.Recordset.EOF Then
               If Not IsNull(FrmContabilizaFacturacion.AdoConsultaFactura.Recordset("ImpuestoxPagar")) Then
                      BuscaCuentaImpuestos = FrmContabilizaFacturacion.AdoConsultaFactura.Recordset("ImpuestoxPagar")
               End If
         End If
         
         
End Function



Function BuscaTasaIvaFactura(CodigoProducto As String) As String
Dim CodigoImpuesto As String

BuscaTasaIvaFactura = 0

CodigoCuentaIva = ""
    '///////////////////////////////////////////////////////////////////////////////////////////////
    '//////////////////////BUSCO LA CUENTA DEL PRODUCTO//////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////
            
            SqlString = "SELECT  * From Productos WHERE (Cod_Productos = '" & CodigoProducto & "')"
            FrmContabilizaFacturacion.AdoConsultaFactura.RecordSource = SqlString
            FrmContabilizaFacturacion.AdoConsultaFactura.Refresh
         If Not FrmContabilizaFacturacion.AdoConsultaFactura.Recordset.EOF Then
               If Not IsNull(FrmContabilizaFacturacion.AdoConsultaFactura.Recordset("Cod_Iva")) Then
                      CodigoImpuesto = FrmContabilizaFacturacion.AdoConsultaFactura.Recordset("Cod_Iva")
               End If
         End If
    
    
            SqlString = "SELECT * From Impuestos WHERE  (Cod_Iva = '" & CodigoImpuesto & "')"
            FrmContabilizaFacturacion.AdoConsultaFactura.RecordSource = SqlString
            FrmContabilizaFacturacion.AdoConsultaFactura.Refresh
         If Not FrmContabilizaFacturacion.AdoConsultaFactura.Recordset.EOF Then
               If Not IsNull(FrmContabilizaFacturacion.AdoConsultaFactura.Recordset("Cod_Iva")) Then
                      CodigoCuentaIva = FrmContabilizaFacturacion.AdoConsultaFactura.Recordset("ImpuestoxPagar")
                      BuscaTasaIvaFactura = FrmContabilizaFacturacion.AdoConsultaFactura.Recordset("Impuesto")
               Else
                     BuscaTasaIvaFactura = 0
               End If
         End If
End Function


Function BuscaTasaIva(CodigoProducto As String) As String
Dim CodigoImpuesto As String

BuscaTasaIva = 0

CodigoCuentaIva = ""
    '///////////////////////////////////////////////////////////////////////////////////////////////
    '//////////////////////BUSCO LA CUENTA DEL PRODUCTO//////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////
            
            SqlString = "SELECT  * From Productos WHERE (Cod_Productos = '" & CodigoProducto & "')"
            FrmContabilizaFacturacion.AdoConsultaFactura.RecordSource = SqlString
            FrmContabilizaFacturacion.AdoConsultaFactura.Refresh
         If Not FrmContabilizaFacturacion.AdoConsultaFactura.Recordset.EOF Then
               If Not IsNull(FrmContabilizaFacturacion.AdoConsultaFactura.Recordset("Cod_Iva")) Then
                      CodigoImpuesto = FrmContabilizaFacturacion.AdoConsultaFactura.Recordset("Cod_Iva")
               End If
         End If
    
    
            SqlString = "SELECT * From Impuestos WHERE  (Cod_Iva = '" & CodigoImpuesto & "')"
            FrmContabilizaFacturacion.AdoConsultaFactura.RecordSource = SqlString
            FrmContabilizaFacturacion.AdoConsultaFactura.Refresh
         If Not FrmContabilizaFacturacion.AdoConsultaFactura.Recordset.EOF Then
               If Not IsNull(FrmContabilizaFacturacion.AdoConsultaFactura.Recordset("Cod_Iva")) Then
                      CodigoCuentaIva = FrmContabilizaFacturacion.AdoConsultaFactura.Recordset("ImpuestoxCobrar")
                      BuscaTasaIva = FrmContabilizaFacturacion.AdoConsultaFactura.Recordset("Impuesto")
               Else
                     BuscaTasaIva = 0
               End If
         End If
End Function

Function BuscaCuentas(CodigoCuenta As String) As Boolean
Dim SqlString As String
  '/////////////////////////////////////////////////////////////////////////////////
  '//////////////////////BUSCO SI EXISTE LA CUENTA////////////////////////////////////
  '////////////////////////////////////////////////////////////////////////////////////
  SqlString = "SELECT * From Cuentas WHERE (CodCuentas = '" & CodigoCuenta & "')"
  MDIPrimero.AdoConsulta.RecordSource = SqlString
  MDIPrimero.AdoConsulta.Refresh
  If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
    BuscaCuentas = True
  Else
    BuscaCuentas = False
  End If
  

End Function

Function TotalMovimientos(FechaTransaccion As Date, NumeroMovimiento As Double) As Double
 Dim Debito As Double, Credito As Double
  MDIPrimero.AdoConsulta.RecordSource = "SELECT SUM(Debito) AS Debito, SUM(Credito) AS Credito From Transacciones WHERE (FechaTransaccion = CONVERT(DATETIME, '" & Format(FechaTransaccion, "yyyy-mm-dd") & "', 102)) AND (NumeroMovimiento = " & NumeroMovimiento & ")"
  MDIPrimero.AdoConsulta.Refresh
  If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
    If Not IsNull(MDIPrimero.AdoConsulta.Recordset("Debito")) Then
      Debito = MDIPrimero.AdoConsulta.Recordset("Debito")
    Else
      Debito = 0
    End If
    
    If Not IsNull(MDIPrimero.AdoConsulta.Recordset("Credito")) Then
      Credito = MDIPrimero.AdoConsulta.Recordset("Credito")
    Else
      Credito = 0
    End If
    TotalMovimientos = Debito - Credito
  End If
 
 
End Function
Function ActualizaConfiguracionReporteResultado()

'Me.CmbUbicacionResultado.AddItem ("***INGRESOS Y VENTAS***")
'Me.CmbUbicacionResultado.AddItem ("")
'Me.CmbUbicacionResultado.AddItem ("")
'Me.CmbUbicacionResultado.AddItem ("")
'Me.CmbUbicacionResultado.AddItem ("")
'Me.CmbUbicacionResultado.AddItem ("***COSTOS Y GASTOS***")
'Me.CmbUbicacionResultado.AddItem ("Compras")
'Me.CmbUbicacionResultado.AddItem ("")
'Me.CmbUbicacionResultado.AddItem ("")
'Me.CmbUbicacionResultado.AddItem ("")
'Me.CmbUbicacionResultado.AddItem ("Acarreo y Fletes")
'Me.CmbUbicacionResultado.AddItem ("Rebajas y Dev S/Compra")
'Me.CmbUbicacionResultado.AddItem ("")
'Me.CmbUbicacionResultado.AddItem ("")
'Me.CmbUbicacionResultado.AddItem ("Gastos")
'Me.CmbUbicacionResultado.AddItem ("")
'Me.CmbUbicacionResultado.AddItem ("")
'Me.CmbUbicacionResultado.AddItem ("")
'Me.CmbUbicacionResultado.AddItem ("")
'Me.CmbUbicacionResultado.AddItem ("")
'Me.CmbUbicacionResultado.AddItem ("Otros Gastos")
'Me.CmbUbicacionResultado.AddItem ("Impuestos Pagados")


    FrmReportes.DtaConsulta.RecordSource = "SELECT * From ConfiguracionReporte"
    FrmReportes.DtaConsulta.Refresh
    If Not FrmReportes.DtaConsulta.Recordset.EOF Then
        If FrmReportes.TxtIngresoVentas.Text <> "" Then
         FrmReportes.DtaConsulta.Recordset("IngresosVentas") = FrmReportes.TxtIngresoVentas.Text
        Else
          FrmReportes.DtaConsulta.Recordset("IngresosVentas") = "Ingresos - Ventas"
        End If
        If FrmReportes.TxtServiciosVentas.Text <> "" Then
         FrmReportes.DtaConsulta.Recordset("ServiciosVentas") = FrmReportes.TxtServiciosVentas.Text
        Else
          FrmReportes.DtaConsulta.Recordset("ServiciosVentas") = "Servicios - Ventas"
        End If
        If FrmReportes.TxtComisionVentas.Text <> "" Then
         FrmReportes.DtaConsulta.Recordset("ComisionVentas") = FrmReportes.TxtComisionVentas.Text
        Else
          FrmReportes.DtaConsulta.Recordset("ComisionVentas") = "Comision - Ventas"
        End If
        If FrmReportes.TxtRebajasyDevolucionesVentas.Text <> "" Then
         FrmReportes.DtaConsulta.Recordset("RebajayDevolucionesVentas") = FrmReportes.TxtRebajasyDevolucionesVentas.Text
        Else
          FrmReportes.DtaConsulta.Recordset("RebajayDevolucionesVentas") = "Rebajas y Dev S/Venta"
        End If
        If FrmReportes.TxtCostodeVentas.Text <> "" Then
         FrmReportes.DtaConsulta.Recordset("CostodeVentas") = FrmReportes.TxtCostodeVentas.Text
        Else
          FrmReportes.DtaConsulta.Recordset("CostodeVentas") = "Costos"
        End If
        If FrmReportes.TxtCostodeProduccion.Text <> "" Then
         FrmReportes.DtaConsulta.Recordset("CostodeProduccion") = FrmReportes.TxtCostodeProduccion.Text
        Else
          FrmReportes.DtaConsulta.Recordset("CostodeProduccion") = "Costos Produccion"
        End If
        If FrmReportes.TxtCostosGeneralesdeProduccion.Text <> "" Then
         FrmReportes.DtaConsulta.Recordset("CostosGeneralesdeProduccion") = FrmReportes.TxtCostosGeneralesdeProduccion.Text
        Else
          FrmReportes.DtaConsulta.Recordset("CostosGeneralesdeProduccion") = "Costos Generales Produccion"
        End If
        If FrmReportes.TxtSueldosyComisiones.Text <> "" Then
         FrmReportes.DtaConsulta.Recordset("SueldosyComisiones") = FrmReportes.TxtSueldosyComisiones.Text
        Else
          FrmReportes.DtaConsulta.Recordset("SueldosyComisiones") = "Sueldos y Comisiones"
        End If
        If FrmReportes.TxtPropaganda.Text <> "" Then
         FrmReportes.DtaConsulta.Recordset("Propaganda") = FrmReportes.TxtPropaganda.Text
        Else
          FrmReportes.DtaConsulta.Recordset("Propaganda") = "Propaganda"
        End If
        If FrmReportes.TxtSueldosAdministrativos.Text <> "" Then
         FrmReportes.DtaConsulta.Recordset("Sueldos") = FrmReportes.TxtSueldosAdministrativos.Text
        Else
          FrmReportes.DtaConsulta.Recordset("Sueldos") = "Sueldos Admon"
        End If
        If FrmReportes.TxtEnergiaElectrica.Text <> "" Then
         FrmReportes.DtaConsulta.Recordset("EnergiaElectrica") = FrmReportes.TxtEnergiaElectrica.Text
        Else
          FrmReportes.DtaConsulta.Recordset("EnergiaElectrica") = "Energia y Agua Potable"
        End If
        If FrmReportes.TxtComisioneGanadas.Text <> "" Then
         FrmReportes.DtaConsulta.Recordset("ComisionesGanadas") = FrmReportes.TxtComisioneGanadas.Text
        Else
          FrmReportes.DtaConsulta.Recordset("ComisionesGanadas") = "Comisiones/Intereses Gandados"
        End If
        If FrmReportes.TxtComisionesPagadas.Text <> "" Then
         FrmReportes.DtaConsulta.Recordset("ComisionesPagadas") = FrmReportes.TxtComisionesPagadas.Text
        Else
          FrmReportes.DtaConsulta.Recordset("ComisionesPagadas") = "Comisiones/Intereses Pagados"
        End If
        If FrmReportes.TxtOtrosIngresos.Text <> "" Then
         FrmReportes.DtaConsulta.Recordset("OtrosIngresosyGastos") = FrmReportes.TxtOtrosIngresos.Text
        Else
          FrmReportes.DtaConsulta.Recordset("OtrosIngresosyGastos") = "Otros Ingresos"
        End If
        If FrmReportes.ChkVentas.Value = xtpChecked Then
           FrmReportes.DtaConsulta.Recordset("AnexosIngresosVentas") = 1
        Else
           FrmReportes.DtaConsulta.Recordset("AnexosIngresosVentas") = 0
        End If
        If FrmReportes.ChkVentasServicios.Value = xtpChecked Then
           FrmReportes.DtaConsulta.Recordset("AnexosServiciosVentas") = 1
        Else
           FrmReportes.DtaConsulta.Recordset("AnexosServiciosVentas") = 0
        End If
        If FrmReportes.ChkComisiones.Value = xtpChecked Then
           FrmReportes.DtaConsulta.Recordset("AnexosComisionVentas") = 1
        Else
           FrmReportes.DtaConsulta.Recordset("AnexosComisionVentas") = 0
        End If
        If FrmReportes.ChkRebajas.Value = xtpChecked Then
           FrmReportes.DtaConsulta.Recordset("AnexosRebajasyDevolucionesVentas") = 1
        Else
           FrmReportes.DtaConsulta.Recordset("AnexosRebajasyDevolucionesVentas") = 0
        End If
        If FrmReportes.ChkCostoVentas.Value = xtpChecked Then
           FrmReportes.DtaConsulta.Recordset("AnexosCostosdeVentas") = 1
        Else
           FrmReportes.DtaConsulta.Recordset("AnexosCostosdeVentas") = 0
        End If
        If FrmReportes.ChkCostoProduccion.Value = xtpChecked Then
           FrmReportes.DtaConsulta.Recordset("AnexosCostosdeProduccion") = 1
        Else
           FrmReportes.DtaConsulta.Recordset("AnexosCostosdeProduccion") = 0
        End If
        If FrmReportes.ChkCostosGeneralesProduccion.Value = xtpChecked Then
           FrmReportes.DtaConsulta.Recordset("AnexosCostosGeneralesdeProduccion") = 1
        Else
           FrmReportes.DtaConsulta.Recordset("AnexosCostosGeneralesdeProduccion") = 0
        End If
        If FrmReportes.ChkSueldosAdmon.Value = xtpChecked Then
           FrmReportes.DtaConsulta.Recordset("AnexosSueldosyComisiones") = 1
        Else
           FrmReportes.DtaConsulta.Recordset("AnexosSueldosyComisiones") = 0
        End If
'        If FrmReportes.ChkSueldos.Value = xtpChecked Then
'           FrmReportes.DtaConsulta.Recordset("AnexosSueldosyComisiones") = 1
'        Else
'           FrmReportes.DtaConsulta.Recordset("AnexosSueldosyComisiones") = 0
'        End If
        If FrmReportes.ChkPropaganda.Value = xtpChecked Then
           FrmReportes.DtaConsulta.Recordset("AnexosPropaganda") = 1
        Else
           FrmReportes.DtaConsulta.Recordset("AnexosPropaganda") = 0
        End If
        If FrmReportes.ChkSueldosAdmon.Value = xtpChecked Then
           FrmReportes.DtaConsulta.Recordset("AnexosSueldos") = 1
        Else
           FrmReportes.DtaConsulta.Recordset("AnexosSueldos") = 0
        End If
        If FrmReportes.ChkEnergiaElectrica.Value = xtpChecked Then
           FrmReportes.DtaConsulta.Recordset("AnexosEnergiaElectrica") = 1
        Else
           FrmReportes.DtaConsulta.Recordset("AnexosEnergiaElectrica") = 0
        End If
        If FrmReportes.ChkComisionesGanadas.Value = xtpChecked Then
           FrmReportes.DtaConsulta.Recordset("AnexosComisionesGanadas") = 1
        Else
           FrmReportes.DtaConsulta.Recordset("AnexosComisionesGanadas") = 0
        End If
        If FrmReportes.ChkComisionesPagadas.Value = xtpChecked Then
           FrmReportes.DtaConsulta.Recordset("AnexosComisionesPagadas") = 1
        Else
           FrmReportes.DtaConsulta.Recordset("AnexosComisionesPagadas") = 0
        End If
        If FrmReportes.ChkOtrosIngresos.Value = xtpChecked Then
           FrmReportes.DtaConsulta.Recordset("AnexosOtrosIngresosyGastos") = 1
        Else
           FrmReportes.DtaConsulta.Recordset("AnexosOtrosIngresosyGastos") = 0
        End If
        
        FrmReportes.DtaConsulta.Recordset.Update

    End If

End Function

Function ActualizaConfiguracionReporte()
    '/////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////EDITO LA CONFIGURACION PARA EL REPORTE///////////////////////
    '//////////////////////////////////////////////////////////////////////////////////////
    
    FrmReportes.DtaConsulta.RecordSource = "SELECT * From ConfiguracionReporte"
    FrmReportes.DtaConsulta.Refresh
    If Not FrmReportes.DtaConsulta.Recordset.EOF Then
     If FrmReportes.TxtCaja.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("Caja") = FrmReportes.TxtCaja.Text
     Else
       FrmReportes.DtaConsulta.Recordset("Caja") = "Caja"
     End If
     
     If FrmReportes.TxtBanco.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("Banco") = FrmReportes.TxtBanco.Text
     Else
       FrmReportes.DtaConsulta.Recordset("Banco") = "Banco"
     End If
   
     If FrmReportes.TxtCtasxCobrar.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("CtasxCobrar") = FrmReportes.TxtCtasxCobrar.Text
     Else
       FrmReportes.DtaConsulta.Recordset("CtasxCobrar") = "Cuentas x Cobrar"
     End If
     
     If FrmReportes.TxtInventario.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("Inventario") = FrmReportes.TxtInventario.Text
     Else
       FrmReportes.DtaConsulta.Recordset("Inventario") = "Inventario"
     End If
    
     If FrmReportes.TxtTerreno.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("Terreno") = FrmReportes.TxtTerreno.Text
     Else
       FrmReportes.DtaConsulta.Recordset("Terreno") = "Terreno y Edificios"
     End If
     
     If FrmReportes.TxtMobiliario.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("Mobiliario") = FrmReportes.TxtMobiliario.Text
     Else
       FrmReportes.DtaConsulta.Recordset("Mobiliario") = "Terreno y Edificios"
     End If
     
     If FrmReportes.TxtEquipoRodante.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("EquipoRodante") = FrmReportes.TxtEquipoRodante.Text
     Else
       FrmReportes.DtaConsulta.Recordset("EquipoRodante") = "Equipo Rodante"
     End If
     
     If FrmReportes.TxtDepreciacion.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("DepAcumulada") = FrmReportes.TxtDepreciacion.Text
     Else
       FrmReportes.DtaConsulta.Recordset("DepAcumulada") = "Depreciacion Acumulada"
     End If
     
     If FrmReportes.TxtPapeleria.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("Papeleria") = FrmReportes.TxtPapeleria.Text
     Else
       FrmReportes.DtaConsulta.Recordset("Papeleria") = "Papeleria y Utiles de Oficina"
     End If
     
     If FrmReportes.TxtPagosAnticipados.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("PagosAnticipados") = FrmReportes.TxtPagosAnticipados.Text
     Else
       FrmReportes.DtaConsulta.Recordset("PagosAnticipados") = "Pagos Anticipados"
     End If
     
     If FrmReportes.TxtProveedores.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("Proveedores") = FrmReportes.TxtProveedores.Text
     Else
       FrmReportes.DtaConsulta.Recordset("Proveedores") = "Proveedores"
     End If
     
     If FrmReportes.TxtOtrosActivos.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("OtrosActivos") = FrmReportes.TxtOtrosActivos.Text
     Else
       FrmReportes.DtaConsulta.Recordset("OtrosActivos") = "Otros Activos"
     End If
     
     If FrmReportes.TxtImpuestosxPagar.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("ImpuestosxPagar") = FrmReportes.TxtImpuestosxPagar.Text
     Else
       FrmReportes.DtaConsulta.Recordset("ImpuestosxPagar") = "Impuestos x Pagar"
     End If
     
     If FrmReportes.TxtDocumentosxPagar.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("DocumentosxPagar") = FrmReportes.TxtDocumentosxPagar.Text
     Else
       FrmReportes.DtaConsulta.Recordset("DocumentosxPagar") = "Documentos x Pagar"
     End If
     
     If FrmReportes.TxtCobrosAnticipados.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("CobroAnticipados") = FrmReportes.TxtCobrosAnticipados.Text
     Else
       FrmReportes.DtaConsulta.Recordset("CobroAnticipados") = "Documentos x Pagar"
     End If
     
     If FrmReportes.TxtPasivosAcumulados.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("PasivosAcumulados") = FrmReportes.TxtPasivosAcumulados.Text
     Else
       FrmReportes.DtaConsulta.Recordset("PasivosAcumulados") = "Pasivos Acumulados"
     End If
     
     If FrmReportes.TxtCtasxPagarLP.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("PagosLP") = FrmReportes.TxtCtasxPagarLP.Text
     Else
       FrmReportes.DtaConsulta.Recordset("PagosLP") = "Cuentas x Pagar LP"
     End If
     
     If FrmReportes.TxtDocumentosxPagarLP.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("DocumentosLP") = FrmReportes.TxtDocumentosxPagarLP.Text
     Else
       FrmReportes.DtaConsulta.Recordset("DocumentosLP") = "Documentos x Pagar LP"
     End If
     
     If FrmReportes.TxtOtrosPasivos.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("OtrosPasivos") = FrmReportes.TxtOtrosPasivos.Text
     Else
       FrmReportes.DtaConsulta.Recordset("OtrosPasivos") = "Otros Pasivos"
     End If
     
     If FrmReportes.TxtAccionesComunes.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("AccionesComunes") = FrmReportes.TxtAccionesComunes.Text
     Else
       FrmReportes.DtaConsulta.Recordset("AccionesComunes") = "Acciones Comunes"
     End If
     
     If FrmReportes.TxtUtilidadAcumulada.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("UtilidadAcumulada") = FrmReportes.TxtUtilidadAcumulada.Text
     Else
       FrmReportes.DtaConsulta.Recordset("UtilidadAcumulada") = "Utilidad Acumulada"
     End If
     
     If FrmReportes.TxtOtrasCtasCapital.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("OtrosCapitales") = FrmReportes.TxtOtrasCtasCapital.Text
     Else
       FrmReportes.DtaConsulta.Recordset("OtrosCapitales") = "Otras Cuentas de Capital"
     End If
     
     If FrmReportes.ChkCaja.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexoCaja") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexoCaja") = 0
     End If
     
     If FrmReportes.ChkBanco.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexoBanco") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexoBanco") = 0
     End If
     
     If FrmReportes.ChkCtasxCob.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexoCtasxCobrar") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexoCtasxCobrar") = 0
     End If
     
     If FrmReportes.ChkInventario.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexoInventario") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexoInventario") = 0
     End If
     
     If FrmReportes.ChkTerreno.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexoTerreno") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexoTerreno") = 0
     End If
     
     If FrmReportes.ChkMobiliario.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexoMobiliario") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexoMobiliario") = 0
     End If
     
     If FrmReportes.ChkEquipoRodante.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexoEquipoRodante") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexoEquipoRodante") = 0
     End If
     
     If FrmReportes.ChkDepreciacionAcum.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexoDepAcumulada") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexoDepAcumulada") = 0
     End If
     
     If FrmReportes.ChkPapeleria.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexoPapeleria") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexoPapeleria") = 0
     End If
     
     If FrmReportes.ChkPagosAnticipados.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosPagosAnticipados") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosPagosAnticipados") = 0
     End If
     
     If FrmReportes.ChkOtrosActivos.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosOtrosActivos") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosOtrosActivos") = 0
     End If
     
     If FrmReportes.ChkProveedores.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosProveedores") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosProveedores") = 0
     End If
     
     If FrmReportes.ChkImpuestosxPagar.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosImpuestosxPagar") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosImpuestosxPagar") = 0
     End If
     
     If FrmReportes.ChkDocumentosxPagar.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosDocumentosxPagar") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosDocumentosxPagar") = 0
     End If
     
     If FrmReportes.ChkCobrosAnticipados.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosCobroAnticipados") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosCobroAnticipados") = 0
     End If
     
     If FrmReportes.ChkPasivosAcum.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosPasivosAcumulados") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosPasivosAcumulados") = 0
     End If
     
     If FrmReportes.ChkCuentasxPagarLP.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosPagosLP") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosPagosLP") = 0
     End If
     
     If FrmReportes.ChkDocumentosxPagLP.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosDocumentosLP") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosDocumentosLP") = 0
     End If
     
     If FrmReportes.ChkOtrosPasivos.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosOtrosPasivos") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosOtrosPasivos") = 0
     End If
     
     If FrmReportes.ChkAccionesComunes.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosAccionesComunes") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosAccionesComunes") = 0
     End If
     
     If FrmReportes.ChkUtilidadAcumulada.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosUtilidadAcumulada") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosUtilidadAcumulada") = 0
     End If
     
     If FrmReportes.ChkUtilidadAcumulada.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosOtrosCapitales") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosOtrosCapitales") = 0
     End If
     
       FrmReportes.DtaConsulta.Recordset.Update
       
    Else
     FrmReportes.DtaConsulta.Recordset.AddNew
     If FrmReportes.TxtCaja.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("Caja") = FrmReportes.TxtCaja.Text
     Else
       FrmReportes.DtaConsulta.Recordset("Caja") = "Caja"
     End If
     
     If FrmReportes.TxtBanco.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("Banco") = FrmReportes.TxtBanco.Text
     Else
       FrmReportes.DtaConsulta.Recordset("Banco") = "Banco"
     End If
   
     If FrmReportes.TxtCtasxCobrar.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("CtasxCobrar") = FrmReportes.TxtCtasxCobrar.Text
     Else
       FrmReportes.DtaConsulta.Recordset("CtasxCobrar") = "Cuentas x Cobrar"
     End If
     
     If FrmReportes.TxtInventario.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("Inventario") = FrmReportes.TxtInventario.Text
     Else
       FrmReportes.DtaConsulta.Recordset("Inventario") = "Inventario"
     End If
    
     If FrmReportes.TxtTerreno.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("Terreno") = FrmReportes.TxtTerreno.Text
     Else
       FrmReportes.DtaConsulta.Recordset("Terreno") = "Terreno y Edificios"
     End If
     
     If FrmReportes.TxtMobiliario.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("Mobiliario") = FrmReportes.TxtMobiliario.Text
     Else
       FrmReportes.DtaConsulta.Recordset("Mobiliario") = "Terreno y Edificios"
     End If
     
     If FrmReportes.TxtEquipoRodante.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("EquipoRodante") = FrmReportes.TxtEquipoRodante.Text
     Else
       FrmReportes.DtaConsulta.Recordset("EquipoRodante") = "Equipo Rodante"
     End If
     
     If FrmReportes.TxtDepreciacion.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("DepAcumulada") = FrmReportes.TxtDepreciacion.Text
     Else
       FrmReportes.DtaConsulta.Recordset("DepAcumulada") = "Depreciacion Acumulada"
     End If
     
     If FrmReportes.TxtPapeleria.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("Papeleria") = FrmReportes.TxtPapeleria.Text
     Else
       FrmReportes.DtaConsulta.Recordset("Papeleria") = "Papeleria y Utiles de Oficina"
     End If
     
     If FrmReportes.TxtPagosAnticipados.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("PagosAnticipados") = FrmReportes.TxtPagosAnticipados.Text
     Else
       FrmReportes.DtaConsulta.Recordset("PagosAnticipados") = "Pagos Anticipados"
     End If
     
     If FrmReportes.TxtProveedores.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("Proveedores") = FrmReportes.TxtProveedores.Text
     Else
       FrmReportes.DtaConsulta.Recordset("Proveedores") = "Proveedores"
     End If
     
     If FrmReportes.TxtOtrosActivos.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("OtrosActivos") = FrmReportes.TxtOtrosActivos.Text
     Else
       FrmReportes.DtaConsulta.Recordset("OtrosActivos") = "Otros Activos"
     End If
     
     If FrmReportes.TxtImpuestosxPagar.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("ImpuestosxPagar") = FrmReportes.TxtImpuestosxPagar.Text
     Else
       FrmReportes.DtaConsulta.Recordset("ImpuestosxPagar") = "Impuestos x Pagar"
     End If
     
     If FrmReportes.TxtDocumentosxPagar.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("DocumentosxPagar") = FrmReportes.TxtDocumentosxPagar.Text
     Else
       FrmReportes.DtaConsulta.Recordset("DocumentosxPagar") = "Documentos x Pagar"
     End If
     
     If FrmReportes.TxtCobrosAnticipados.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("CobroAnticipados") = FrmReportes.TxtCobrosAnticipados.Text
     Else
       FrmReportes.DtaConsulta.Recordset("CobroAnticipados") = "Documentos x Pagar"
     End If
     
     If FrmReportes.TxtPasivosAcumulados.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("PasivosAcumulados") = FrmReportes.TxtPasivosAcumulados.Text
     Else
       FrmReportes.DtaConsulta.Recordset("PasivosAcumulados") = "Pasivos Acumulados"
     End If
     
     If FrmReportes.TxtCtasxPagarLP.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("PagosLP") = FrmReportes.TxtCtasxPagarLP.Text
     Else
       FrmReportes.DtaConsulta.Recordset("PagosLP") = "Cuentas x Pagar LP"
     End If
     
     If FrmReportes.TxtDocumentosxPagarLP.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("DocumentosLP") = FrmReportes.TxtDocumentosxPagarLP.Text
     Else
       FrmReportes.DtaConsulta.Recordset("DocumentosLP") = "Documentos x Pagar LP"
     End If
     
     If FrmReportes.TxtOtrosPasivos.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("OtrosPasivos") = FrmReportes.TxtOtrosPasivos.Text
     Else
       FrmReportes.DtaConsulta.Recordset("OtrosPasivos") = "Otros Pasivos"
     End If
     
     If FrmReportes.TxtAccionesComunes.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("AccionesComunes") = FrmReportes.TxtAccionesComunes.Text
     Else
       FrmReportes.DtaConsulta.Recordset("AccionesComunes") = "Acciones Comunes"
     End If
     
     If FrmReportes.TxtUtilidadAcumulada.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("UtilidadAcumulada") = FrmReportes.TxtUtilidadAcumulada.Text
     Else
       FrmReportes.DtaConsulta.Recordset("UtilidadAcumulada") = "Utilidad Acumulada"
     End If
     
     If FrmReportes.TxtOtrasCtasCapital.Text <> "" Then
      FrmReportes.DtaConsulta.Recordset("OtrosCapitales") = FrmReportes.TxtOtrasCtasCapital.Text
     Else
       FrmReportes.DtaConsulta.Recordset("OtrosCapitales") = "Otras Cuentas de Capital"
     End If
     
     If FrmReportes.ChkCaja.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexoCaja") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexoCaja") = 0
     End If
     
     If FrmReportes.ChkBanco.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexoBanco") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexoBanco") = 0
     End If
     
     If FrmReportes.ChkCtasxCob.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexoCtasxCobrar") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexoCtasxCobrar") = 0
     End If
     
     If FrmReportes.ChkInventario.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexoInventario") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexoInventario") = 0
     End If
     
     If FrmReportes.ChkTerreno.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexoTerreno") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexoTerreno") = 0
     End If
     
     If FrmReportes.ChkMobiliario.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexoMobiliario") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexoMobiliario") = 0
     End If
     
     If FrmReportes.ChkEquipoRodante.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexoEquipoRodante") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexoEquipoRodante") = 0
     End If
     
     If FrmReportes.ChkDepreciacionAcum.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexoDepAcumulada") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexoDepAcumulada") = 0
     End If
     
     If FrmReportes.ChkPapeleria.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexoPapeleria") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexoPapeleria") = 0
     End If
     
     If FrmReportes.ChkPagosAnticipados.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosPagosAnticipados") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosPagosAnticipados") = 0
     End If
     
     If FrmReportes.ChkOtrosActivos.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosOtrosActivos") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosOtrosActivos") = 0
     End If
     
     If FrmReportes.ChkProveedores.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosProveedores") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosProveedores") = 0
     End If
     
     If FrmReportes.ChkImpuestosxPagar.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosImpuestosxPagar") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosImpuestosxPagar") = 0
     End If
     
     If FrmReportes.ChkDocumentosxPagar.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosDocumentosxPagar") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosDocumentosxPagar") = 0
     End If
     
     If FrmReportes.ChkCobrosAnticipados.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosCobroAnticipados") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosCobroAnticipados") = 0
     End If
     
     If FrmReportes.ChkPasivosAcum.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosPasivosAcumulados") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosPasivosAcumulados") = 0
     End If
     
     If FrmReportes.ChkCuentasxPagarLP.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosPagosLP") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosPagosLP") = 0
     End If
     
     If FrmReportes.ChkDocumentosxPagLP.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosDocumentosLP") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosDocumentosLP") = 0
     End If
     
     If FrmReportes.ChkOtrosPasivos.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosOtrosPasivos") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosOtrosPasivos") = 0
     End If
     
     If FrmReportes.ChkAccionesComunes.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosAccionesComunes") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosAccionesComunes") = 0
     End If
     
     If FrmReportes.ChkUtilidadAcumulada.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosUtilidadAcumulada") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosUtilidadAcumulada") = 0
     End If
     
     If FrmReportes.ChkUtilidadAcumulada.Value = xtpChecked Then
      FrmReportes.DtaConsulta.Recordset("AnexosOtrosCapitales") = 1
     Else
       FrmReportes.DtaConsulta.Recordset("AnexosOtrosCapitales") = 0
     End If
     
     
       FrmReportes.DtaConsulta.Recordset.Update
    End If
End Function


Function ReporteResumenAnexos(FechaIni As Date, FechaFin As Date)
Dim Fechas1 As String, Fechas2 As String, Orden As Integer, Sql As String, i As Double
Dim UltimoOrden As Integer, RegIngresos  As Integer, PrimReg As Integer, UltReg As Integer
Dim Utilidad As Double, Utilidad2 As Double, Utilidad3 As Double, RegTCostosOper As Integer
Dim Decrementador As Integer, TotalActivoCirculante As Double, TotalActivoFijo As Double, TotalActivoDiferido As Double
Dim TotalPasivoCirculante As Double, TotalPasivoFijo As Double, TotalPasivoDiferido As Double, TotalCapitalSocial As Double
Dim RegInicioCostosOperativos As Integer 'variable que me guarda el orden donde se encuentra el registro donde comienzan los costos operativos
Dim RegTotalCostosOperativos As Integer 'variable que me guarda el orden donde se encuentra el registro de total de costos operativos
Dim Totalingresos As Double, TotalCostoVentas As Double, TotalGastosAdmon As Double, TotalGastos As Double
Dim TotalGastoVentas As Double, TotalIngresosFinancieros As Double, TotalOtrosIngresos As Double, TotalOtrosGastos As Double
Dim TotalUtilidadBruta As Double, TotalImpuestos As Double, TotalUtilidadNeta As Double, Fecha1 As String, Fecha2 As String
Dim TotalCompras As Double, TotalInventarioInicial As Double, TotalInventarioFinal As Double
Dim TotalAcarreo As Double, TotalRebajaVentas As Double, TotalDisponible As Double, TotalGastosR As Double, TotalCosto As Double
Dim TotalSalidas As Double, TotalGastoOperacion As Double, TotalPasivo As Double, TotalCapital As Double
Dim TotalCostos As Double, ListaActivos As Variant, TotalInventario As Double, TotalCuentaxCobrar As Double
Dim TotalCuentasxPagar As Double, TotalActivos As Double, UtilidadBrutas As Double, UtilidadNetas As Double
Dim ListaMeses As Variant, CantRegistros As Double, ComboIni As Double, ComboFin As Double, TotalCostoFijo As Double, TotalGastoFijo As Double
Dim mes As Double


    ArepBalancePersonalizado.Logo.Picture = LoadPicture(RutaLogo)
    ArepBalancePersonalizado.LblMoneda.Caption = "Expresado en " & FrmReportes.CmbMoneda.Text
    ArepBalancePersonalizado.LblEmpresa = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
    ArepBalancePersonalizado.LblEmpresa1 = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
    ArepBalancePersonalizado.LblEmpresa2 = "RUC: " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
    ArepBalancePersonalizado.LblFechaFin = Format(FechaFin, "dd/mm/yyyy")
    ArepBalancePersonalizado.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
    ArepBalancePersonalizado.LblFechaIni = Format(FechaIni, "dd/mm/yyyy")

    '////////////////////RESULTADOS DE ACTIVO CIRCULANTE//////////////////////////////
    TotalActivoCirculante = 0
    SaldosPersonalizados ("Cajas")
    ArepBalancePersonalizado.LblCajas.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoCirculante = TotalActivoCirculante + ResultadoPersonalizado
    SaldosPersonalizados ("Bancos")
    ArepBalancePersonalizado.LblBancos.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoCirculante = TotalActivoCirculante + ResultadoPersonalizado
    SaldosPersonalizados ("Inventario")
    ArepBalancePersonalizado.LblInventario.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoCirculante = TotalActivoCirculante + ResultadoPersonalizado
    SaldosPersonalizados ("Cuentas x Cobrar")
    ArepBalancePersonalizado.LblCtasCobrar.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoCirculante = TotalActivoCirculante + ResultadoPersonalizado
    ArepBalancePersonalizado.LblTotalActivoCirculante.Caption = Format(TotalActivoCirculante, "##,##0.00")
    
    
    '//////////////////RESULTADOS DE ACTIVO FIJO//////////////////////////////////////////
    SaldosPersonalizados ("Terreno y Edificios")
    ArepBalancePersonalizado.LblTerreno.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoFijo = TotalActivoFijo + ResultadoPersonalizado
    SaldosPersonalizados ("Mobiliario y Equipo de Oficina")
    ArepBalancePersonalizado.LblMobiliario.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoFijo = TotalActivoFijo + ResultadoPersonalizado
    SaldosPersonalizados ("Equipo Rodante")
    ArepBalancePersonalizado.LblEquipoRodante.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoFijo = TotalActivoFijo + ResultadoPersonalizado
    SaldosPersonalizados ("Depreciacion Acumulada")
    ArepBalancePersonalizado.LblDepreciacion.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoFijo = TotalActivoFijo + ResultadoPersonalizado
    ArepBalancePersonalizado.LblTotalActivoFijo.Caption = Format(TotalActivoFijo, "##,##0.00")
    
    '//////////////////RESULTADOS DE ACTIVO DIFERIDO//////////////////////////////////////////
    SaldosPersonalizados ("Papeleria y Utiles de Oficina")
    ArepBalancePersonalizado.LblPapeleria.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoDiferido = TotalActivoDiferido + ResultadoPersonalizado
    SaldosPersonalizados ("Pagos Anticipados")
    ArepBalancePersonalizado.LblAnticipos.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoDiferido = TotalActivoDiferido + ResultadoPersonalizado
    SaldosPersonalizados ("Otros Activos")
    ArepBalancePersonalizado.LblOtrosActivos.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoDiferido = TotalActivoDiferido + ResultadoPersonalizado
    ArepBalancePersonalizado.LblTotalActivoDiferido.Caption = Format(TotalActivoDiferido, "##,##0.00")
    
    '////////////////PASIVO CIRCULANTE//////////////////////////////////////////////////////
    SaldosPersonalizados ("Proveedores")
    ArepBalancePersonalizado.LblProveedores.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalPasivoCirculante = TotalPasivoCirculante + ResultadoPersonalizado
    SaldosPersonalizados ("Impuestos x Pagar")
    ArepBalancePersonalizado.LblImpuestosPagar.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalPasivoCirculante = TotalPasivoCirculante + ResultadoPersonalizado
    SaldosPersonalizados ("Documentos x Pagar CP")
    ArepBalancePersonalizado.LblDocumentosPagar.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalPasivoCirculante = TotalPasivoCirculante + ResultadoPersonalizado
    SaldosPersonalizados ("Cobros Anticipados")
    ArepBalancePersonalizado.LblCobrosAnticipados.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalPasivoCirculante = TotalPasivoCirculante + ResultadoPersonalizado
    SaldosPersonalizados ("Pasivos Acumulados")
    ArepBalancePersonalizado.LblPasivosAcumulados.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalPasivoCirculante = TotalPasivoCirculante + ResultadoPersonalizado
    ArepBalancePersonalizado.LblTotalPasivoCirculante.Caption = Format(TotalPasivoCirculante, "##,##0.00")
    
    '////////////////PASIVO FIJO//////////////////////////////////////////////////////
    SaldosPersonalizados ("Cuentas x Pagar LP")
    ArepBalancePersonalizado.LblCuentasPagarLP.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalPasivoFijo = TotalPasivoFijo + ResultadoPersonalizado
    SaldosPersonalizados ("Documentos x Pagar LP")
    ArepBalancePersonalizado.LblDocumentosPagarLP.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalPasivoFijo = TotalPasivoFijo + ResultadoPersonalizado
    ArepBalancePersonalizado.LblTotalPasivoFijo.Caption = Format(TotalPasivoFijo, "##,##0.00")
    
    '//////////////PASIVO DIFERIDO///////////////////////////////////////////////
    SaldosPersonalizados ("Otros Pasivos")
    ArepBalancePersonalizado.LblOtrosPasivos.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalPasivoDiferido = TotalPasivoDiferido + ResultadoPersonalizado
    ArepBalancePersonalizado.LblTotalPasivoDiferido.Caption = Format(TotalPasivoDiferido, "##,##0.00")
    
     '//////////////CAPITAL///////////////////////////////////////////////
    SaldosPersonalizados ("Acciones Comunes")
    ArepBalancePersonalizado.LblAccionesComunes.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalCapitalSocial = TotalCapitalSocial + ResultadoPersonalizado
    SaldosPersonalizados ("Utilidades Acumuladas")
    ArepBalancePersonalizado.LblUtilidadesAcumuladas.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalCapitalSocial = TotalCapitalSocial + ResultadoPersonalizado
    SaldosPersonalizados ("Otras Ctas de Capital")
    ArepBalancePersonalizado.LblOtrasCuentasCapital.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalCapitalSocial = TotalCapitalSocial + ResultadoPersonalizado
    SaldosPersonalizados ("Resultado Periodo")
    ArepBalancePersonalizado.LblResultadoPeriodo.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalCapitalSocial = TotalCapitalSocial + ResultadoPersonalizado
    ArepBalancePersonalizado.LblTotalCapital.Caption = Format(TotalCapitalSocial, "##,##0.00")
    
    ArepBalancePersonalizado.LblTotalPasivomasCapital.Caption = Format(TotalPasivoCirculante + TotalPasivoFijo + TotalPasivoDiferido + TotalCapitalSocial, "##,##0.00")
    ArepBalancePersonalizado.LblTotalActivo.Caption = Format(TotalActivoCirculante + TotalActivoFijo + TotalActivoDiferido, "##,##0.00")
    
    

    
End Function


Function AnexosReporteResumen(FechaIni As Date, FechaFin As Date)
  Dim EncabezadoConsulta As String, Condiciones As String
  Dim Sql As String
  
  '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  '/////////////////////////////GUARDO EL ENZABEZADO DE LA CONSULTA/////////////////////////////////////////////////////
  '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  EncabezadoConsulta = "SELECT Reportes.Descripcion AS Descripcion, Reportes.Debe1 AS Debe1, Reportes.Haber1 AS Haber1, Reportes.Debe2 AS Debe2,Reportes.Haber2 AS Haber2, Reportes.Debe3 AS Debe3, Reportes.Haber3 AS Haber3, Reportes.KeyGrupo AS KeyGrupo,Reportes.KeyGrupoSuperior AS KeyGrupoSuperior, Reportes.KeyGrupoCuenta AS KeyGrupoCuenta, Reportes.Nivel AS Nivel, Reportes.Orden AS Orden,Reportes.CodCuentas AS CodCuentas, Cuentas.DescripcionCuentas, Reportes.Ubicacion " & _
                       "FROM Reportes INNER JOIN Cuentas ON Reportes.KeyGrupo = Cuentas.CodCuentas  "
  
  Condiciones = ""
  
  If FrmReportes.ChkCaja.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Cajas') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Cajas')"
    End If
  End If
  
  If FrmReportes.ChkBanco.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Bancos') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Bancos')"
    End If
  End If
  
  If FrmReportes.ChkInventario.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Inventario') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Inventario')"
    End If
  End If
  
  If FrmReportes.ChkCtasxCob.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Cuentas x Cobrar') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Cuentas x Cobrar')"
    End If
  End If
  
  If FrmReportes.ChkTerreno.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Terreno y Edificios') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Terreno y Edificios')"
    End If
  End If
  
  If FrmReportes.ChkMobiliario.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Mobiliario y Equipo de Oficina') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Mobiliario y Equipo de Oficina')"
    End If
  End If
  
  If FrmReportes.ChkEquipoRodante.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Equipo Rodante') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Equipo Rodante')"
    End If
  End If
  
  If FrmReportes.ChkDepreciacionAcum.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Depreciacion Acumulada') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Depreciacion Acumulada')"
    End If
  End If

  If FrmReportes.ChkPapeleria.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Papeleria y Utiles de Oficina') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Papeleria y Utiles de Oficina')"
    End If
  End If
  
   If FrmReportes.ChkPagosAnticipados.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Pagos Anticipados') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Pagos Anticipados')"
    End If
   End If
   
   If FrmReportes.ChkOtrosActivos.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Otros Activos') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Otros Activos')"
    End If
   End If
   
   If FrmReportes.ChkProveedores.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Proveedores') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Proveedores')"
    End If
   End If

   If FrmReportes.ChkImpuestosxPagar.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Impuestos x Pagar') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Impuestos x Pagar')"
    End If
   End If
   
   If FrmReportes.ChkDocumentosxPagar.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Documentos x Pagar CP') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Documentos x Pagar CP')"
    End If
   End If
   
   If FrmReportes.ChkCobrosAnticipados.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Cobros Anticipados') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Cobros Anticipados')"
    End If
   End If
   
   If FrmReportes.ChkPasivosAcum.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Pasivos Acumulados') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Pasivos Acumulados')"
    End If
   End If
   
   If FrmReportes.ChkCuentasxPagarLP.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Cuentas x Pagar LP') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Cuentas x Pagar LP')"
    End If
   End If
   
   If FrmReportes.ChkDocumentosxPagLP.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Documentos x Pagar LP') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Documentos x Pagar LP')"
    End If
   End If
   
   If FrmReportes.ChkOtrosPasivos.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Otros Pasivos') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Otros Pasivos')"
    End If
   End If
   
   If FrmReportes.ChkAccionesComunes.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Acciones Comunes') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Acciones Comunes')"
    End If
   End If
   
   If FrmReportes.ChkUtilidadAcumulada.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Utilidades Acumuladas') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Utilidades Acumuladas')"
    End If
   End If
   
   If FrmReportes.ChkOtrasCtasCapital.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Otras Ctas de Capital') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Otras Ctas de Capital')"
    End If
   End If
   
If Condiciones <> "" Then
 Sql = EncabezadoConsulta & Condiciones & " ORDER BY Reportes.Orden"
    ArepAnexosBalances.Logo.Picture = LoadPicture(RutaLogo)
'    ArepAnexosBalances.LblMoneda.Caption = "Expresado en " & FrmReportes.CmbMoneda.Text
    ArepAnexosBalances.LblEmpresa = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
    ArepAnexosBalances.LblEmpresa1 = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
    ArepAnexosBalances.LblEmpresa2 = "RUC: " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
    ArepAnexosBalances.LblFechaFin = Format(FechaFin, "dd/mm/yyyy")
    ArepAnexosBalances.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
    ArepAnexosBalances.LblFechaIni = Format(FechaIni, "dd/mm/yyyy")
    ArepAnexosBalances.DataControl1.ConnectionString = ConexionReporte
    ArepAnexosBalances.DataControl1.Source = Sql
    ArepAnexosBalances.Show 1
End If


 
End Function

Function ConfiguracionReportesBalance()
'///////////////////////////////////////////CARGO LA CONFIGURACION  /////////////////////////
    FrmReportes.DtaConsulta.RecordSource = "SELECT * From ConfiguracionReporte"
    FrmReportes.DtaConsulta.Refresh
    If Not FrmReportes.DtaConsulta.Recordset.EOF Then
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("Caja")) Then
      FrmReportes.TxtCaja.Text = FrmReportes.DtaConsulta.Recordset("Caja")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("Banco")) Then
      FrmReportes.TxtBanco.Text = FrmReportes.DtaConsulta.Recordset("Banco")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("CtasxCobrar")) Then
     FrmReportes.TxtCtasxCobrar.Text = FrmReportes.DtaConsulta.Recordset("CtasxCobrar")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("Inventario")) Then
     FrmReportes.TxtInventario.Text = FrmReportes.DtaConsulta.Recordset("Inventario")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("Terreno")) Then
     FrmReportes.TxtTerreno.Text = FrmReportes.DtaConsulta.Recordset("Terreno")
     End If
     If Not IsNull(FrmReportes.TxtMobiliario.Text = FrmReportes.DtaConsulta.Recordset("Mobiliario")) Then
     FrmReportes.TxtMobiliario.Text = FrmReportes.DtaConsulta.Recordset("Mobiliario")
     End If
     If Not IsNull(FrmReportes.TxtEquipoRodante.Text = FrmReportes.DtaConsulta.Recordset("EquipoRodante")) Then
     FrmReportes.TxtEquipoRodante.Text = FrmReportes.DtaConsulta.Recordset("EquipoRodante")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("DepAcumulada")) Then
     FrmReportes.TxtDepreciacion.Text = FrmReportes.DtaConsulta.Recordset("DepAcumulada")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("Papeleria")) Then
     FrmReportes.TxtPapeleria.Text = FrmReportes.DtaConsulta.Recordset("Papeleria")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("PagosAnticipados")) Then
     FrmReportes.TxtPagosAnticipados.Text = FrmReportes.DtaConsulta.Recordset("PagosAnticipados")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("OtrosActivos")) Then
     FrmReportes.TxtOtrosActivos.Text = FrmReportes.DtaConsulta.Recordset("OtrosActivos")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("Proveedores")) Then
     FrmReportes.TxtProveedores.Text = FrmReportes.DtaConsulta.Recordset("Proveedores")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("ImpuestosxPagar")) Then
     FrmReportes.TxtImpuestosxPagar.Text = FrmReportes.DtaConsulta.Recordset("ImpuestosxPagar")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("DocumentosxPagar")) Then
     FrmReportes.TxtDocumentosxPagar.Text = FrmReportes.DtaConsulta.Recordset("DocumentosxPagar")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("CobroAnticipados")) Then
     FrmReportes.TxtCobrosAnticipados.Text = FrmReportes.DtaConsulta.Recordset("CobroAnticipados")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("PasivosAcumulados")) Then
     FrmReportes.TxtPasivosAcumulados.Text = FrmReportes.DtaConsulta.Recordset("PasivosAcumulados")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("PagosLP")) Then
     FrmReportes.TxtCtasxPagarLP.Text = FrmReportes.DtaConsulta.Recordset("PagosLP")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("DocumentosLP")) Then
     FrmReportes.TxtDocumentosxPagarLP.Text = FrmReportes.DtaConsulta.Recordset("DocumentosLP")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("OtrosPasivos")) Then
     FrmReportes.TxtOtrosPasivos.Text = FrmReportes.DtaConsulta.Recordset("OtrosPasivos")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AccionesComunes")) Then
     FrmReportes.TxtAccionesComunes.Text = FrmReportes.DtaConsulta.Recordset("AccionesComunes")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("UtilidadAcumulada")) Then
     FrmReportes.TxtUtilidadAcumulada.Text = FrmReportes.DtaConsulta.Recordset("UtilidadAcumulada")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("OtrosCapitales")) Then
     FrmReportes.TxtOtrasCtasCapital.Text = FrmReportes.DtaConsulta.Recordset("OtrosCapitales")
     End If

     If Not IsNull(FrmReportes.DtaConsulta.Recordset("OtrosCapitales")) Then
     FrmReportes.TxtOtrasCtasCapital.Text = FrmReportes.DtaConsulta.Recordset("OtrosCapitales")
     End If
     
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexoCaja")) = True Then
      If FrmReportes.DtaConsulta.Recordset("AnexoCaja") = True Then
       FrmReportes.ChkCaja.Value = xtpChecked
      End If
     End If
     
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexoBanco")) = True Then
       If FrmReportes.DtaConsulta.Recordset("AnexoBanco") = True Then
       FrmReportes.ChkBanco.Value = xtpChecked
       End If
     End If
     
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexoCtasxCobrar")) = True Then
      If FrmReportes.DtaConsulta.Recordset("AnexoCtasxCobrar") = True Then
       FrmReportes.ChkCtasxCob.Value = xtpChecked
      End If
     End If
    
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexoInventario")) = True Then
      If FrmReportes.DtaConsulta.Recordset("AnexoInventario") = True Then
       FrmReportes.ChkInventario.Value = xtpChecked
      End If
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexoTerreno")) = True Then
      If FrmReportes.DtaConsulta.Recordset("AnexoTerreno") = True Then
       FrmReportes.ChkTerreno.Value = xtpChecked
      End If
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexoMobiliario")) = True Then
      If FrmReportes.DtaConsulta.Recordset("AnexoMobiliario") = True Then
       FrmReportes.ChkMobiliario.Value = xtpChecked
      End If
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexoEquipoRodante")) = True Then
      If FrmReportes.DtaConsulta.Recordset("AnexoEquipoRodante") = True Then
       FrmReportes.ChkEquipoRodante.Value = xtpChecked
      End If
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexoDepAcumulada")) = True Then
       If FrmReportes.DtaConsulta.Recordset("AnexoDepAcumulada") = True Then
       FrmReportes.ChkDepreciacionAcum.Value = xtpChecked
       End If
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexoPapeleria")) = True Then
      If FrmReportes.DtaConsulta.Recordset("AnexoPapeleria") = True Then
       FrmReportes.ChkPapeleria.Value = xtpChecked
      End If
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexosPagosAnticipados")) = True Then
       If FrmReportes.DtaConsulta.Recordset("AnexosPagosAnticipados") = True Then
       FrmReportes.ChkPagosAnticipados.Value = xtpChecked
       End If
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexosOtrosActivos")) = True Then
       If FrmReportes.DtaConsulta.Recordset("AnexosOtrosActivos") = True Then
       FrmReportes.ChkOtrosActivos.Value = xtpChecked
       End If
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexosProveedores")) = True Then
       If FrmReportes.DtaConsulta.Recordset("AnexosProveedores") = True Then
       FrmReportes.ChkProveedores.Value = xtpChecked
       End If
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexosImpuestosxPagar")) = True Then
       If FrmReportes.DtaConsulta.Recordset("AnexosImpuestosxPagar") = True Then
       FrmReportes.ChkImpuestosxPagar.Value = xtpChecked
       End If
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexosDocumentosxPagar")) = True Then
       If FrmReportes.DtaConsulta.Recordset("AnexosDocumentosxPagar") = True Then
       FrmReportes.ChkDocumentosxPagar.Value = xtpChecked
       End If
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexosCobroAnticipados")) = True Then
       If FrmReportes.DtaConsulta.Recordset("AnexosCobroAnticipados") = True Then
       FrmReportes.ChkCobrosAnticipados.Value = xtpChecked
       End If
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexosPasivosAcumulados")) = True Then
       If FrmReportes.DtaConsulta.Recordset("AnexosPasivosAcumulados") = True Then
       FrmReportes.ChkPasivosAcum.Value = xtpChecked
       End If
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexosPagosLP")) = True Then
       If FrmReportes.DtaConsulta.Recordset("AnexosPagosLP") = True Then
       FrmReportes.ChkCuentasxPagarLP.Value = xtpChecked
       End If
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexosDocumentosLP")) = True Then
       If FrmReportes.DtaConsulta.Recordset("AnexosDocumentosLP") = True Then
       FrmReportes.ChkDocumentosxPagLP.Value = xtpChecked
       End If
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexosOtrosPasivos")) = True Then
       If FrmReportes.DtaConsulta.Recordset("AnexosOtrosPasivos") = True Then
       FrmReportes.ChkOtrosPasivos.Value = xtpChecked
       End If
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexosAccionesComunes")) = True Then
       If FrmReportes.DtaConsulta.Recordset("AnexosAccionesComunes") = True Then
       FrmReportes.ChkAccionesComunes.Value = xtpChecked
       End If
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexosUtilidadAcumulada")) = True Then
       If FrmReportes.DtaConsulta.Recordset("AnexosUtilidadAcumulada") = True Then
       FrmReportes.ChkUtilidadAcumulada.Value = xtpChecked
       End If
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("AnexosOtrosCapitales")) = True Then
       If FrmReportes.DtaConsulta.Recordset("AnexosOtrosCapitales") = True Then
       FrmReportes.ChkOtrasCtasCapital.Value = xtpChecked
       End If
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("IngresosVentas")) Then
     FrmReportes.TxtIngresoVentas.Text = FrmReportes.DtaConsulta.Recordset("IngresosVentas")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("ServiciosVentas")) Then
      FrmReportes.TxtServiciosVentas.Text = FrmReportes.DtaConsulta.Recordset("ServiciosVentas")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("ComisionVentas")) Then
      FrmReportes.TxtComisionVentas.Text = FrmReportes.DtaConsulta.Recordset("ComisionVentas")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("RebajayDevolucionesVentas")) Then
      FrmReportes.TxtRebajasyDevolucionesVentas.Text = FrmReportes.DtaConsulta.Recordset("RebajayDevolucionesVentas")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("CostodeVentas")) Then
      FrmReportes.TxtCostodeVentas.Text = FrmReportes.DtaConsulta.Recordset("CostodeVentas")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("CostodeProduccion")) Then
      FrmReportes.TxtCostodeProduccion.Text = FrmReportes.DtaConsulta.Recordset("CostodeProduccion")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("CostosGeneralesdeProduccion")) Then
      FrmReportes.TxtCostosGeneralesdeProduccion.Text = FrmReportes.DtaConsulta.Recordset("CostosGeneralesdeProduccion")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("SueldosyComisiones")) Then
      FrmReportes.TxtSueldosyComisiones.Text = FrmReportes.DtaConsulta.Recordset("SueldosyComisiones")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("Propaganda")) Then
      FrmReportes.TxtPropaganda.Text = FrmReportes.DtaConsulta.Recordset("Propaganda")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("Sueldos")) Then
      FrmReportes.TxtSueldosAdministrativos.Text = FrmReportes.DtaConsulta.Recordset("Sueldos")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("EnergiaElectrica")) Then
      FrmReportes.TxtEnergiaElectrica.Text = FrmReportes.DtaConsulta.Recordset("EnergiaElectrica")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("ComisionesGanadas")) Then
      FrmReportes.TxtComisioneGanadas.Text = FrmReportes.DtaConsulta.Recordset("ComisionesGanadas")
     End If
     If Not IsNull(FrmReportes.DtaConsulta.Recordset("ComisionesPagadas")) Then
      FrmReportes.TxtComisionesPagadas.Text = FrmReportes.DtaConsulta.Recordset("ComisionesPagadas")
     End If
    End If
End Function

Public Function SumasDebitos(NumeroMovimiento As Double, NPeriodo As Double)
 Dim Sql As String, Debito As Double, Credito As Double
  Sql = "SELECT  FechaTransaccion AS FechaTransaccion, TCambio AS TCambio, Debito AS Debito, Credito AS Credito From Transacciones Where (NumeroMovimiento = " & NumeroMovimiento & ") And (NPeriodo = " & NPeriodo & ")"
  MDIPrimero.AdoConsulta.RecordSource = Sql
  MDIPrimero.AdoConsulta.Refresh
  Do While Not MDIPrimero.AdoConsulta.Recordset.EOF
   Debito = Debito + MDIPrimero.AdoConsulta.Recordset("Debito")
   Credito = Credito + MDIPrimero.AdoConsulta.Recordset("Credito")
   MDIPrimero.AdoConsulta.Recordset.MoveNext
  Loop

  FrmReparar.TxtCredito.Text = Credito
  FrmReparar.TxtDebito.Text = Debito
  FrmReparar.TxtDiferencia.Text = Format(Debito - Credito, "##,##0.0000")
End Function


Public Sub CargaADODC(TablaMaestra As String, ByRef nombreADODC As Adodc, xsinonimo, nombreCombo As String, apliTrim As String, conex As String, f As Form, sqlOrd As String, Optional estado As String)
    nombreADODC.ConnectionString = Conexion
    
    If TablaMaestra = "CatCombus" Then
            Sql = "select (descripcombus)," & TablaMaestra & ".idcombus from " & TablaMaestra & " " & sqlOrd
    End If
    
    If TablaMaestra = "CataVH" Then
            Sql = "select (descricpcion)," & TablaMaestra & ".idvh from " & TablaMaestra & " " & sqlOrd
    End If
    
    If TablaMaestra = "_Sede" Then
            Sql = "select (Descripcion)," & TablaMaestra & ".IdSede from " & TablaMaestra & " where activo=" & xsinonimo & " " & sqlOrd
    End If
    
    If TablaMaestra = "ControlSeccionFinca" Then
        Sql = "select IdReg, NoSeccion from ControlSeccionFinca where IdFinca=" & FrmPlanificaActividad.cmdfinca.BoundText & " and  anoplanta=" & Trim(FrmPlanificaActividad.txtplantado.Text) & " "
    End If
    
     If TablaMaestra = "TipoNomina" Then
        Sql = "select CodTipoNomina, Nomina from TipoNomina where activa='True'"
    End If
        
    If TablaMaestra = "_Finca" Then
        If apliTrim = "" Then
            Sql = "select ((Finca))," & TablaMaestra & ".IdFinca from " & TablaMaestra & " where IdSede=" & FrmRegSeccion.sede.BoundText & " " & sqlOrd
        Else
            If apliTrim = "2" Then
                Sql = "select ((Finca))," & TablaMaestra & ".IdFinca from " & TablaMaestra & " where IdSede=" & FrmPlanificaActividad.cmdsede.BoundText & " " & sqlOrd
            Else
                Sql = "select ((Finca))," & TablaMaestra & ".IdFinca from " & TablaMaestra & "  " & sqlOrd
            End If
        End If
    End If
    
    If TablaMaestra = "_Plantacion" Then
            Sql = "select (Plantacion)," & TablaMaestra & ".IdPlantacion from " & TablaMaestra & "  " & sqlOrd
    End If
    If TablaMaestra = "DatosEmpresa" Then
            Sql = "select NombreEmpresa,ConexionSistemaContable, " & TablaMaestra & ".Numero from " & TablaMaestra & " where  Numero=2 " & sqlOrd
    End If
    If TablaMaestra = "Cargo" Then
            Sql = "select Cargo, " & TablaMaestra & ".CodCargo from " & TablaMaestra & "  " & sqlOrd
    End If
    If TablaMaestra = "tab_units" Then
            Sql = "select name, " & TablaMaestra & ".id_unit from " & TablaMaestra & "  " & sqlOrd
    End If
    If TablaMaestra = "Productos" Then
            Sql = "select Descripcion_Producto, " & TablaMaestra & ".Cod_Productos from " & TablaMaestra & " WHERE Cod_Cuenta_Inventario LIKE '" & wcodbodega & "%'" & sqlOrd
    End If
    If TablaMaestra = "Tareas" Then
        If wregNuSede = 0 Then
            If IsNumeric(FrmPlanificaActividad.cmdsede.BoundText) Then
                wregNuSede = FrmPlanificaActividad.cmdsede.BoundText
            End If
        End If
            Sql = "select CodigoTarea as No_Tarea, (LOWER (Nombre_Tarea)) as Nombre_Tarea , QuienPaga,UnidMedida, DH from  " & TablaMaestra & " where CodigoTarea like '" & wregNuSede & "%'"
    End If
    
    If TablaMaestra = "Empleado" Then
            Sql = "select (nombre1 +' '+ nombre2 +' '+ apellido1 +' '+ apellido2) as nombrecompleto, Empleado.CodEmpleado from Empleado where (activo=1 or activo='True') "
    End If

    On Local Error Resume Next
    err = 0
    If TablaMaestra = "Productos" Or TablaMaestra = "Tareas" Then
        nombreADODC.ConnectionString = ConexionInventario
    Else
        If TablaMaestra = "CataVH" Or TablaMaestra = "CatCombus" Then
            nombreADODC.ConnectionString = Conexion
        Else
            nombreADODC.ConnectionString = Conexion
        End If
    End If
    nombreADODC.RecordSource = Sql
    nombreADODC.Refresh
    If err <> 0 Then
        nombreADODC.ConnectionString = Conexion
        nombreADODC.RecordSource = Sql
        nombreADODC.Refresh
    End If
    
'    Set nombreADODC.Recordset.ActiveConnection = Nothing
End Sub
Public Function tienepermiso(usu As Integer, aplica As String, oper As String) As Boolean
Set rsa = Nothing
Sql = "select permitido from mapermisos where cci_rif=" & usu & " and aplicacion='" & aplica & "' and operacion='" & oper & "' "
rsa.Open Sql, Conexion, adOpenForwardOnly, adLockReadOnly
If rsa.EOF = True Then
    tienepermiso = False
Else
    If rsa!permitido = 1 Then
        tienepermiso = True
    Else
        tienepermiso = False
    End If
End If
End Function

Public Sub CargaADODCConta(TablaMaestra As String, ByRef nombreADODC As Adodc, xsinonimo, nombreCombo As String, apliTrim As String, conex As String, f As Form, sqlOrd As String, Optional estado As String)
nombreADODC.ConnectionString = Conexion
If TablaMaestra = "Oficinas" Then
    Sql = "select (Descripcion)," & TablaMaestra & ".Idreg from " & TablaMaestra & "  " & sqlOrd
End If

If TablaMaestra = "ResponsablesAreas" Then
    Sql = "select (NombreResponsable)," & TablaMaestra & ".Idreg from " & TablaMaestra & "  " & sqlOrd
End If

On Local Error Resume Next
    err = 0
    nombreADODC.ConnectionString = Conexion
    nombreADODC.RecordSource = Sql
    nombreADODC.Refresh
    If err <> 0 Then
        nombreADODC.ConnectionString = Conexion
        nombreADODC.RecordSource = Sql
        nombreADODC.Refresh
    End If
End Sub

Function EsCedulaValida(cedula As String) As Boolean
   EsCedulaValida = False
   If Len(cedula) = 16 Then
      If IsNumeric(Mid(cedula, 1, 3)) And Mid(cedula, 4, 1) = "-" And IsNumeric(Mid(cedula, 5, 6)) And Mid(cedula, 11, 1) = "-" And IsNumeric(Mid(cedula, 12, 4)) And Mid(cedula, 16, 1) = ObtenerLetra(cedula) Then
         'Asi: MaxTeen. 2010-08-18 10:53:24 am. Validando correctamente la cadena de seis dígitos para saber si es de tipo fecha válida
         'If IsDate(Mid(cedula, 5, 2) & "/" & Mid(cedula, 7, 2) & "/" & Mid(cedula, 9, 2)) Then
         If EsFecha_SeisDigitos(Mid(cedula, 5, 2) & Mid(cedula, 7, 2) & Mid(cedula, 9, 2)) = True Then
            EsCedulaValida = True
         End If
      End If
   End If
End Function

Public Function EsFecha_SeisDigitos(cadfecha As String) As Boolean
'Asi: MaxTeen. 2010-08-18 10:50:30 am. Nueva función para validar si el parámetro es una fecha válida
Dim dd As Integer, MM As Integer
Dim aa As Integer, bisiesto As Boolean
'''
EsFecha_SeisDigitos = True
bisiesto = False
'cadfecha = Mid(Trim(cadfecha),  5, 6)
'''
dd = Val(Mid(cadfecha, 1, 2))
MM = Val(Mid(cadfecha, 3, 2))
aa = Val(Mid(cadfecha, 5, 2))
If aa Mod 2 = 0 Then
    If aa Mod 4 = 0 Then bisiesto = True
End If
If MM < 1 Or MM > 12 Then
    EsFecha_SeisDigitos = False
    Exit Function
Else
    Select Case MM
        Case 1: Case 3: Case 5: Case 7: Case 8: Case 10: Case 12
            If dd < 1 Or dd > 31 Then EsFecha_SeisDigitos = False: Exit Function
        Case 2
            If bisiesto Then
                If dd < 1 Or dd > 29 Then EsFecha_SeisDigitos = False: Exit Function
            Else
                If dd < 1 Or dd > 28 Then EsFecha_SeisDigitos = False: Exit Function
            End If
        Case Else
            If dd < 1 Or dd > 30 Then EsFecha_SeisDigitos = False: Exit Function
    End Select
End If
End Function

Public Function ArchivoActual(d As String) As String
Dim AactL As Integer, ind%, pleca%
'Dim ArchActual
d = Trim(d)
AactL = Len(Trim(d))
For ind = 1 To AactL
    If Mid(d, ind, 1) = "\" Then
        pleca = ind
    End If
Next ind
ArchivoActual = Mid(d, pleca + 1, AactL)
End Function

Function Mod2(cadena, num) As Long
   Dim Texto As String
   Dim num1, num2 As Long

   Texto = ""
   num1 = 0
   num2 = num

   For i = 1 To Len(cadena)
      Texto = Texto & Mid(cadena, i, 1)
      num1 = Val(Texto)
      Res = num1 Mod num2
      Texto = Str(Res)
   Next

   Mod2 = Res
End Function

Function ObtenerLetra(cedula As String) As String
   Dim cadena As String
   Dim posicion As Long
   cadena = "ABCDEFGHJKLMNPQRSTUVWXY"
   posicion = Mod2(cedula, 23) + 1
   ObtenerLetra = Mid(cadena, posicion, 1)
End Function

Public Function DameNombreEmpleado(codempl) As String
Set rsa1 = Nothing
If codempl <> 0 Then
    Sql = "select nombre1, nombre2, apellido1, apellido2, codtiponomina, codgrupo, codcargo, SueldoPeriodo, dolarizado from empleado where codempleado=" & codempl & ""
Else
    Sql = "select nombre1, nombre2, apellido1, apellido2, codtiponomina, codgrupo, codcargo, SueldoPeriodo,dolarizado from empleado where codempleado=" & FrmFormulario.AdoHist.Recordset!idsoli & ""
End If
rsa1.Open Sql, Conexion, adOpenForwardOnly, adLockOptimistic
If rsa1.EOF = True Then
    Set rsa1 = Nothing
    Sql = "select nombre1, nombre2, apellido1, apellido2, codtiponomina, codgrupo, codcargo, SueldoPeriodo,dolarizado from empleado where codempleado1=" & codempl & ""
    rsa1.Open Sql, Conexion, adOpenForwardOnly, adLockOptimistic
End If
DameNombreEmpleado = rsa1!Nombre1 & " " & rsa1!Nombre2 & " " & rsa1!Apellido1 & " " & rsa1!Apellido2
idnomina = rsa1!CodTipoNomina
salanterior = rsa1!SueldoPeriodo
If IsNull(rsa1!CodGrupo) Then
    idgrupo = 0
Else
    idgrupo = rsa1!CodGrupo
    grupo = NombreGrupo(rsa1!CodGrupo)
End If

If IsNull(rsa1!codcargo) Then
    idcargo = 0
Else
    idcargo = rsa1!codcargo
'    nomcargo = NombreCargo(rsa1!codcargo)
End If
Nomina = NombreNominaC(rsa1!CodTipoNomina)
isaldolar = rsa1!dolarizado
End Function
Public Function NombreNominaC(Cod As String) As String
Set rsa2 = Nothing
Sql = "select nomina from dbo.TipoNomina where codtiponomina = '" & Cod & "'"
rsa2.Open Sql, Conexion, adOpenForwardOnly, adLockOptimistic
NombreNominaC = rsa2!Nomina
End Function

Public Function NombreGrupo(Cod As String) As String
Set rsa2 = Nothing
Sql = "select grupo from grupo where codgrupo= '" & Cod & "'"

rsa2.Open Sql, Conexion, adOpenForwardOnly, adLockOptimistic
NombreGrupo = rsa2!grupo
End Function

Public Function DepartamentoID(Descripcion As String) As String
Set rsa2 = Nothing
Sql = "SELECT  * From Oficinas WHERE (Descripcion = '" & Descripcion & " ')"
rsa2.Open Sql, Conexion, adOpenForwardOnly, adLockOptimistic
DepartamentoID = rsa2!idreg
End Function
Public Function ResponsableID(Descripcion As String) As String
Set rsa2 = Nothing
Sql = "SELECT  * From ResponsablesAreas WHERE  (NombreResponsable = '" & Descripcion & " ')"
rsa2.Open Sql, Conexion, adOpenForwardOnly, adLockOptimistic
ResponsableID = rsa2!idreg
End Function
