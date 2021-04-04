VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepResultadoResumen 
   Caption         =   "Estado de Resultado Resumen"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   19080
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   33655
   _ExtentY        =   19368
   SectionData     =   "ArepResultadoResumen.dsx":0000
End
Attribute VB_Name = "ArepResultadoResumen"
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
Dim Fechas1 As String, Fechas2 As String, Orden As Integer, SQL As String, i As Double
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
Dim Mes As Double, R As Variant, TipoReporte As String


    FrmReportes.DtaReportes.Refresh
    Me.LblMoneda.Caption = "Expresado en " & FrmReportes.CmbMoneda.Text
    Me.Logo.Picture = LoadPicture(RutaLogo)
    Me.LblEmpresa = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
    Me.LblEmpresa1 = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
    Me.LblEmpresa2 = "RUC: " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
    Me.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
    Me.LblFechaFin = FechaFin
    Me.LblFechaIni = FechaIni
'    me.LblAcumulado.Caption = "SALDO HASTA " & FechaFin
'    me.LblActual.Caption = "ACTIVIDAD PERIODO"
'    me.LblAnterior.Caption = "SALDO ANTES " & FechaIni
    Me.LblBalance.Caption = FrmReportes.TxtTipoReporte.Text
    
    Fecha1 = Format(FechaIni, "yyyy-mm-dd")
    Fecha2 = Format(FechaFin, "yyyy-mm-dd")
    

  '////////////////////////////////BUSCO EL INVENTARIO INCIAL ////////////////////////////////////////////////////////////
        SQL = "SELECT  MAX(Transacciones.CodCuentas) AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Debito) - SUM(Transacciones.TCambio * Transacciones.Credito) AS Saldo, MAX(Cuentas.KeyGrupo) AS KeyGrupo, MAX(Cuentas.DescripcionGrupo)AS DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta  " & _
              "FROM    Grupos INNER JOIN Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo  " & _
              "WHERE  (Transacciones.FechaTransaccion < CONVERT(DATETIME, '" & Fecha1 & "', 102)) " & _
              "GROUP BY Cuentas.TipoCuenta HAVING (Cuentas.TipoCuenta = 'Inventario') ORDER BY MAX(Cuentas.KeyGrupo)"
   FrmReportes.DtaConsulta.RecordSource = SQL
   FrmReportes.DtaConsulta.Refresh
   If Not FrmReportes.DtaConsulta.Recordset.EOF Then
     TotalInventarioInicial = FrmReportes.DtaConsulta.Recordset("Saldo")
   Else
     TotalInventarioInicial = 0
   End If
    

    '////////////////////////////////////BUSCO LAS COMPRAS DEL PERIODO///////////////////////////////////////////////////////
    SQL = "SELECT MAX(Transacciones.CodCuentas) AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Debito) - SUM(Transacciones.TCambio * Transacciones.Credito) AS Saldo, MAX(Cuentas.KeyGrupo) AS KeyGrupo, MAX(Cuentas.DescripcionGrupo) AS DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta  " & _
          "FROM  Grupos INNER JOIN Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo " & _
          "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102))  " & _
          "GROUP BY Cuentas.TipoCuenta HAVING  (Cuentas.TipoCuenta = 'Inventario')ORDER BY MAX(Cuentas.KeyGrupo)"
    FrmReportes.DtaConsulta.RecordSource = SQL
    FrmReportes.DtaConsulta.Refresh
    If Not FrmReportes.DtaConsulta.Recordset.EOF Then
         TotalCompras = FrmReportes.DtaConsulta.Recordset("Debito")
         TotalSalidas = FrmReportes.DtaConsulta.Recordset("Credito")
    Else
         TotalCompras = 0
         TotalSalidas = 0
    End If
    
   '////////////////////////////////BUSCO EL INVENTARIO FINAL ////////////////////////////////////////////////////////////
   SQL = "SELECT  MAX(Transacciones.CodCuentas) AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Debito) - SUM(Transacciones.TCambio * Transacciones.Credito) AS Saldo, MAX(Cuentas.KeyGrupo) AS KeyGrupo, MAX(Cuentas.DescripcionGrupo)AS DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta  " & _
         "FROM    Grupos INNER JOIN Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo  " & _
         "WHERE  (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
         "GROUP BY Cuentas.TipoCuenta HAVING (Cuentas.TipoCuenta = 'Inventario') ORDER BY MAX(Cuentas.KeyGrupo)"
   FrmReportes.DtaConsulta.RecordSource = SQL
   FrmReportes.DtaConsulta.Refresh
   If Not FrmReportes.DtaConsulta.Recordset.EOF Then
     TotalInventarioFinal = FrmReportes.DtaConsulta.Recordset("Saldo")
   Else
     TotalInventarioFinal = 0
   End If
    
    
    '////////////////////RESULTADOS DE INGRESOS//////////////////////////////
    If FrmReportes.OptAcumulado.Value = True Then
        Totalingresos = 0
        IngresosVentas = 0
        ServiciosVentas = 0
        ComisionVentas = 0
                
        SaldosPersonalizados ("Ingresos - Ventas")
        IngresosVentas = ResultadoPersonalizado
        Me.LblVentas.Caption = Format(IngresosVentas, "##,##0.00")
        SaldosPersonalizados ("Servicios - Ventas")
        ServiciosVentas = ResultadoPersonalizado
        Me.LblVentasServicios.Caption = Format(ServiciosVentas, "##,##0.00")
        SaldosPersonalizados ("Comision - Ventas")
        ComisionVentas = ResultadoPersonalizado
        Me.LblComisiones.Caption = Format(ComisionVentas, "##,##0.00")
        Totalingresos = IngresosVentas + ServiciosVentas + ComisionVentas
'        Me.LblVentas.Caption = Format(Totalingresos, "##,##0.00")
        Totalingresos = Totalingresos + ResultadoPersonalizado
        SaldosPersonalizados ("Rebajas y Dev S/Venta")
        Me.LblRebajas.Caption = Format(ResultadoPersonalizado, "##,##0.00")
        Totalingresos = Totalingresos - ResultadoPersonalizado
        Me.LblTotalIngreso.Caption = Format(Totalingresos, "##,##0.00")
    ElseIf FrmReportes.OptPeriodo.Value = True Then
         Totalingresos = 0
        IngresosVentas = 0
        ServiciosVentas = 0
        ComisionVentas = 0
        SaldosPersonalizados ("Ingresos - Ventas")
        IngresosVentas = ResultadoPersonalizado
        Me.LblVentas.Caption = Format(IngresosVentas, "##,##0.00")
        SaldosPersonalizados ("Servicios - Ventas")
        ServiciosVentas = ResultadoPersonalizado
        Me.LblVentasServicios.Caption = Format(ServiciosVentas, "##,##0.00")
        SaldosPersonalizados ("Comision - Ventas")
        ComisionVentas = ResultadoPersonalizado
        Me.LblComisiones.Caption = Format(ComisionVentas, "##,##0.00")
        Totalingresos = IngresosVentas + ServiciosVentas + ComisionVentas
'        Me.LblVentas.Caption = Format(Totalingresos, "##,##0.00")
        Totalingresos = Totalingresos + ResultadoPersonalizadoPeriodo
        SaldosPersonalizados ("Rebajas y Dev S/Venta")
        Me.LblRebajas.Caption = Format(ResultadoPersonalizadoPeriodo, "##,##0.00")
        Totalingresos = Totalingresos - ResultadoPersonalizadoPeriodo
        Me.LblTotalIngreso.Caption = Format(Totalingresos, "##,##0.00")
     End If
    
    '////////////////////RESULTADO DE COSTOS/////////////////////////////////////
     If FrmReportes.OptAcumulado.Value = True Then
        TotalCostoVentas = 0
        Me.LblInventarioInicial.Caption = Format(TotalInventarioInicial, "##,##0.00")
        Me.LblCompras.Caption = Format(TotalCompras, "##,##0.00")
        SaldosPersonalizados ("Acarreo y Fletes")
        Me.LblAcarreos.Caption = Format(ResultadoPersonalizado, "##,##0.00")
        TotalAcarreo = ResultadoPersonalizado
        SaldosPersonalizados ("Rebajas y Dev S/Compra")
        Me.LblRebajasCompra.Caption = Format(ResultadoPersonalizado, "##,##0.00")
        TotalRebajaVentas = ResultadoPersonalizado
        
        SaldosPersonalizados ("Costos")
        TotalCosto = ResultadoPersonalizado
        SaldosPersonalizados ("Costos Produccion")
        TotalCosto = TotalCosto + ResultadoPersonalizado
        SaldosPersonalizados ("Costos Generales Produccion")
        TotalCosto = TotalCosto + ResultadoPersonalizado
        
        Me.LblCostosProductos.Caption = Format(TotalCosto, "##,##0.00")
        TotalDisponible = TotalInventarioInicial + TotalCompras + TotalAcarreo + TotalCosto - TotalRebajaVentas
        Me.LblDisponible.Caption = Format(TotalDisponible, "##,##0.00")
        Me.LblInventarioFinal.Caption = Format(TotalInventarioFinal, "##,##0.00")
        Me.LblSalidasInventarios.Caption = Format(TotalSalidas, "##,##0.00")
        TotalCostoVentas = TotalDisponible - TotalInventarioFinal - TotalSalidas
        Me.LblTotalCostoVentas.Caption = Format(TotalCostoVentas, "##,##0.00")
        TotalUtilidadBruta = Totalingresos - TotalCostoVentas
        Me.LblUtilidad.Caption = Format(TotalUtilidadBruta, "##,##0.00")
     ElseIf FrmReportes.OptPeriodo.Value = True Then
        TotalCostoVentas = 0
        Me.LblInventarioInicial.Caption = Format(TotalInventarioInicial, "##,##0.00")
        Me.LblCompras.Caption = Format(TotalCompras, "##,##0.00")
        SaldosPersonalizados ("Acarreo y Fletes")
        Me.LblAcarreos.Caption = Format(ResultadoPersonalizadoPeriodo, "##,##0.00")
        TotalAcarreo = ResultadoPersonalizadoPeriodo
        SaldosPersonalizados ("Rebajas y Dev S/Compra")
        Me.LblRebajasCompra.Caption = Format(ResultadoPersonalizadoPeriodo, "##,##0.00")
        TotalRebajaVentas = ResultadoPersonalizadoPeriodo
        SaldosPersonalizados ("Costos")
        TotalCosto = ResultadoPersonalizadoPeriodo
        Me.LblCostosProductos.Caption = Format(TotalCosto, "##,##0.00")
        TotalDisponible = TotalInventarioInicial + TotalCompras + TotalAcarreo + TotalCosto - TotalRebajaVentas
        Me.LblDisponible.Caption = Format(TotalDisponible, "##,##0.00")
        Me.LblInventarioFinal.Caption = Format(TotalInventarioFinal, "##,##0.00")
        Me.LblSalidasInventarios.Caption = Format(TotalSalidas, "##,##0.00")
        TotalCostoVentas = TotalDisponible - TotalInventarioFinal - TotalSalidas
        Me.LblTotalCostoVentas.Caption = Format(TotalCostoVentas, "##,##0.00")
        TotalUtilidadBruta = Totalingresos - TotalCostoVentas
        Me.LblUtilidad.Caption = Format(TotalUtilidadBruta, "##,##0.00")
     
     End If
     
     If FrmReportes.OptAcumulado.Value = True Then
        TotalGastoVentas = 0
        SaldosPersonalizados ("Sueldos y Comisiones")
        Me.LblSueldosVentas.Caption = Format(ResultadoPersonalizado, "##,##0.00")
        TotalGastoVentas = TotalGastoVentas + ResultadoPersonalizado
        SaldosPersonalizados ("Propaganda")
        Me.LblPropaganda.Caption = Format(ResultadoPersonalizado, "##,##0.00")
        TotalGastoVentas = TotalGastoVentas + ResultadoPersonalizado
        Me.LblTotalGatosVentas.Caption = Format(TotalGastoVentas, "##,##0.00")
     ElseIf FrmReportes.OptPeriodo.Value = True Then
        TotalGastoVentas = 0
        SaldosPersonalizados ("Sueldos y Comisiones")
        Me.LblSueldosVentas.Caption = Format(ResultadoPersonalizadoPeriodo, "##,##0.00")
        TotalGastoVentas = TotalGastoVentas + ResultadoPersonalizadoPeriodo
        SaldosPersonalizados ("Propaganda")
        Me.LblPropaganda.Caption = Format(ResultadoPersonalizadoPeriodo, "##,##0.00")
        TotalGastoVentas = TotalGastoVentas + ResultadoPersonalizadoPeriodo
        Me.LblTotalGatosVentas.Caption = Format(TotalGastoVentas, "##,##0.00")
     End If
     
    If FrmReportes.OptAcumulado.Value = True Then
        TotalGastosAdmon = 0
        SaldosPersonalizados ("Sueldos Admon")
        Me.LblSueldosAdmon.Caption = Format(ResultadoPersonalizado, "##,##0.00")
        TotalGastosAdmon = TotalGastosAdmon + ResultadoPersonalizado
        SaldosPersonalizados ("Gastos")
        TotalGastosR = ResultadoPersonalizado
        TotalGastosAdmon = TotalGastosAdmon + TotalGastosR
        SaldosPersonalizados ("Energia y Agua Potable")
        Me.LblEnergia.Caption = Format(ResultadoPersonalizado + TotalGastosR, "##,##0.00")
        TotalGastosAdmon = TotalGastosAdmon + ResultadoPersonalizado
        Me.LblTotalGastosAdmon.Caption = Format(TotalGastosAdmon, "##,##0.00")
    ElseIf FrmReportes.OptPeriodo.Value = True Then
        TotalGastosAdmon = 0
        SaldosPersonalizados ("Sueldos Admon")
        Me.LblSueldosAdmon.Caption = Format(ResultadoPersonalizadoPeriodo, "##,##0.00")
        TotalGastosAdmon = TotalGastosAdmon + ResultadoPersonalizadoPeriodo
        SaldosPersonalizados ("Gastos")
        TotalGastosR = ResultadoPersonalizadoPeriodo
        TotalGastosAdmon = TotalGastosAdmon + TotalGastosR
        SaldosPersonalizados ("Energia y Agua Potable")
        Me.LblEnergia.Caption = Format(ResultadoPersonalizadoPeriodo + TotalGastosR, "##,##0.00")
        TotalGastosAdmon = TotalGastosAdmon + ResultadoPersonalizadoPeriodo
        Me.LblTotalGastosAdmon.Caption = Format(TotalGastosAdmon, "##,##0.00")
    
    End If
    

        TotalGastoOperacion = 0
        TotalGastoOperacion = TotalGastosAdmon + TotalGastoVentas
        Me.LblTotalGastoOperacion.Caption = Format(TotalGastoOperacion, "##,##0.00")

     If FrmReportes.OptAcumulado.Value = True Then
        TotalIngresosFinancieros = 0
        SaldosPersonalizados ("Comisiones/Intereses Gandados")
        Me.LblComisionesGanadas.Caption = Format(ResultadoPersonalizado, "##,##0.00")
        TotalIngresosFinancieros = TotalIngresosFinancieros + ResultadoPersonalizado
        SaldosPersonalizados ("Comisiones/Intereses Pagados")
        Me.LblComisionesPagadas.Caption = Format(ResultadoPersonalizado, "##,##0.00")
        TotalIngresosFinancieros = TotalIngresosFinancieros + ResultadoPersonalizado
        Me.LblTotalIngresosFinancieros.Caption = Format(TotalIngresosFinancieros, "##,##0.00")
     Else
        TotalIngresosFinancieros = 0
        SaldosPersonalizados ("Comisiones/Intereses Gandados")
        Me.LblComisionesGanadas.Caption = Format(ResultadoPersonalizadoPeriodo, "##,##0.00")
        TotalIngresosFinancieros = TotalIngresosFinancieros + ResultadoPersonalizadoPeriodo
        SaldosPersonalizados ("Comisiones/Intereses Pagados")
        Me.LblComisionesPagadas.Caption = Format(ResultadoPersonalizadoPeriodo, "##,##0.00")
        TotalIngresosFinancieros = TotalIngresosFinancieros + ResultadoPersonalizadoPeriodo
        Me.LblTotalIngresosFinancieros.Caption = Format(TotalIngresosFinancieros, "##,##0.00")

     End If
     
     If FrmReportes.OptAcumulado.Value = True Then
        TotalOtrosIngresos = 0
        SaldosPersonalizados ("Otros Ingresos")
        TotalOtrosIngresos = ResultadoPersonalizado + TotalOtrosIngresos
     Else
        TotalOtrosIngresos = 0
        SaldosPersonalizados ("Otros Ingresos")
        TotalOtrosIngresos = ResultadoPersonalizadoPeriodo + TotalOtrosIngresos
     End If
     
     If FrmReportes.OptAcumulado.Value = True Then
        TotalOtrosGastos = 0
        SaldosPersonalizados ("Otros Gastos")
        TotalOtrosGastos = TotalOtrosGastos + ResultadoPersonalizado
        Me.LblOtrosIngresos.Caption = Format(TotalOtrosIngresos - TotalOtrosGastos, "##,##0.00")
     Else
        TotalOtrosGastos = 0
        SaldosPersonalizados ("Otros Gastos")
        TotalOtrosGastos = TotalOtrosGastos + ResultadoPersonalizadoPeriodo
        Me.LblOtrosIngresos.Caption = Format(TotalOtrosIngresos - TotalOtrosGastos, "##,##0.00")

     End If
     
     If FrmReportes.OptAcumulado.Value = True Then
        TotalImpuestos = 0
        SaldosPersonalizados ("Impuestos Pagados")
        TotalImpuestos = TotalImpuestos + ResultadoPersonalizado
        Me.LblImpuestos.Caption = Format(TotalImpuestos, "##,##0.00")
     Else
        TotalImpuestos = 0
        SaldosPersonalizados ("Impuestos Pagados")
        TotalImpuestos = TotalImpuestos + ResultadoPersonalizadoPeriodo
        Me.LblImpuestos.Caption = Format(TotalImpuestos, "##,##0.00")
     
     End If
     
     '/////////////////////////////////////BUSCO LA UTILIDAD DEL PERIODO ////////////////////////
     MDIPrimero.AdoConsulta.RecordSource = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden, CodCuentas, Ubicacion, Orden AS Expr1, Haber3 - Debe3 AS Resultado From Reportes WHERE (KeyGrupo = 'RP') ORDER BY Expr1"
     MDIPrimero.AdoConsulta.Refresh
     If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
       TotalUtilidadNeta = MDIPrimero.AdoConsulta.Recordset("Resultado")
  
     End If
'    TotalUtilidadNeta = TotalUtilidadBruta - (TotalGastoVentas + TotalGastosAdmon + TotalIngresosFinancieros + TotalOtrosGastos) + TotalOtrosIngresos - TotalImpuestos
    Me.LblUtilidadNeta.Caption = Format(TotalUtilidadNeta, "##,##0.00")
    
    TipoReporte = FrmReportes.TxtTipoReporte.Text
    
    If TipoReporte = "ESTADO DE RESULTADO RESUMEN ANEXOS" Then
        FrmReportes.DtaConsulta.RecordSource = "SELECT * From ConfiguracionReporte"
        FrmReportes.DtaConsulta.Refresh
        If Not FrmReportes.DtaConsulta.Recordset.EOF Then
         Me.LblNombreVentas.Caption = FrmReportes.DtaConsulta.Recordset("IngresosVentas")
         Me.LblNombreVentaServicios.Caption = FrmReportes.DtaConsulta.Recordset("ServiciosVentas")
         Me.LblNombreComisiones.Caption = FrmReportes.DtaConsulta.Recordset("ComisionVentas")
         Me.LblNombreRebajas.Caption = FrmReportes.DtaConsulta.Recordset("RebajayDevolucionesVentas")
         Me.LblNombreCostoVentas.Caption = FrmReportes.DtaConsulta.Recordset("CostodeVentas")
         Me.LblNombreTotalCostoVentas.Caption = "Total " & FrmReportes.DtaConsulta.Recordset("CostodeVentas")
         Me.LblNombreCostoProduccion.Caption = FrmReportes.DtaConsulta.Recordset("CostodeProduccion")
         Me.LblNombreSueldosComisiones.Caption = FrmReportes.DtaConsulta.Recordset("SueldosyComisiones")
         Me.LblNombrePropaganda.Caption = FrmReportes.DtaConsulta.Recordset("Propaganda")
         Me.LblNombreSueldosAdministracion.Caption = FrmReportes.DtaConsulta.Recordset("Sueldos")
         Me.LblNombreEnergia.Caption = FrmReportes.DtaConsulta.Recordset("EnergiaElectrica")
         Me.LblNombreComisionesGanadas.Caption = FrmReportes.DtaConsulta.Recordset("ComisionesGanadas")
         Me.LblNombreComisionPagada.Caption = FrmReportes.DtaConsulta.Recordset("ComisionesPagadas")
         Me.LblNombreOtrosIngresos.Caption = FrmReportes.DtaConsulta.Recordset("OtrosIngresosyGastos")
        End If
    End If
           
    FrmReportes.LblProgreso.Caption = ""
    FrmReportes.osProgress1.Visible = False
End Sub

