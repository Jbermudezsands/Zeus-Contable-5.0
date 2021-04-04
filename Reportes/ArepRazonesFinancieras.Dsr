VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepRazonesFinancieras 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   19080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   33655
   _ExtentY        =   20267
   SectionData     =   "ArepRazonesFinancieras.dsx":0000
End
Attribute VB_Name = "ArepRazonesFinancieras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ActiveReport_ReportStart()
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


            NumeroPeriodo1 = FrmReportes.CmbIni.Text
            NumeroPeriodo2 = FrmReportes.CmbFin.Text
            
            If FrmReportes.Option8 = True Then
             NumeroTabla = 1
            ElseIf FrmReportes.Option7 = True Then
              NumeroTabla = 2
            ElseIf FrmReportes.Option6 = True Then
              NumeroTabla = 3
            End If
            
              FrmReportes.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) Between " & NumeroPeriodo1 & " And " & NumeroPeriodo2 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
              FrmReportes.DtaConsulta.Refresh
              If FrmReportes.DtaConsulta.Recordset.RecordCount = 0 Then
                MsgBox "Hay un problema con los Periodos, por Favor Revise los Periodos para buscar alguna anomalía", vbCritical
                Exit Sub
              End If
               FrmReportes.DtaConsulta.Recordset.MoveLast
               i = FrmReportes.DtaConsulta.Recordset.RecordCount
               FrmReportes.DtaConsulta.Recordset.MoveFirst
               
               FrmReportes.osProgress1.Visible = True
               FrmReportes.osProgress1.Value = 0
               FrmReportes.osProgress1.Max = i
'               frmreportes.blProgreso.Caption = "Procesando las fechas"
              Do While Not FrmReportes.DtaConsulta.Recordset.EOF
        
        
                If i = 1 Then
                  FechaIni = "01/" & Month(FrmReportes.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
                  NumFecha1 = FechaIni
                  FechaFin = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
                  NumFecha2 = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
                Else
        
                 If NumeroPeriodo1 = FrmReportes.DtaConsulta.Recordset("Periodo") Then
                  FechaIni = "01/" & Month(FrmReportes.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
                  NumFecha1 = FechaIni
                 ElseIf NumeroPeriodo2 = FrmReportes.DtaConsulta.Recordset("Periodo") Then
                  FechaFin = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
                  NumFecha2 = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
                 End If
                End If
                
                FrmReportes.osProgress1.Value = FrmReportes.osProgress1.Value + 1
                DoEvents
                
                FrmReportes.DtaConsulta.Recordset.MoveNext
              Loop
              
              Fecha1 = Format(FechaIni, "yyyy-mm-dd")
              Fecha2 = Format(FechaFin, "yyyy-mm-dd")

                
                '//////////////////////////////////////////////////////////////////////////////////
                '/////////////////SUMO LOS ACTIVOS CIRCULANTES/////////////////////////////////////
                '//////////////////////////////////////////////////////////////////////////////////
               ListaActivos = Array("Caja", "Bancos", "Inventario", "Cuentas x Cobrar", "Papeleria - Utiles", "Otros Activos")
               TotalActivoCirculante = 0
               i = 0
               FrmReportes.osProgress1.Value = 0
               FrmReportes.osProgress1.Max = 6
               For i = 0 To 5
                TotalActivoCirculante = SaldosRazonesDebitos(Fecha2, ListaActivos(i)) + TotalActivoCirculante
                FrmReportes.osProgress1.Value = FrmReportes.osProgress1.Value + 1
                DoEvents
               Next
               
               '//////////////////////////SUMO LOS ACTIVOS FIJOS////////////////////////////////////////////////////////
                TotalActivoFijo = 0
                TotalActivoFijo = SaldosRazonesDebitos(Fecha2, "Activo Fijo")
                
                '///////////////////////SUMO LOS ACTIVOS//////////////////////////////////////////
                TotalActivos = TotalActivoCirculante + TotalActivoFijo
                
                '//////////////////////////SUMO LOS INVENTARIOS////////////////////////////////////////////////////////
                TotalInventario = 0
                TotalInventario = SaldosRazonesDebitos(Fecha2, "Inventario")
                
                '//////////////////////////SUMO LAS CUENTAS X COBRAR////////////////////////////////////////////////////////
                TotalCuentaxCobrar = 0
                TotalCuentaxCobrar = SaldosRazonesDebitos(Fecha2, "Cuentas x Cobrar")
                
                '/////////////////////TOTAL PASIVOS//////////////////////////////////////////////////////////////////////////
                i = 0
                TotalPasivo = 0
                ListaActivos = Array("Cuentas x Pagar", "Otros Pasivos", "Pasivo")
                
                FrmReportes.osProgress1.Value = 0
                FrmReportes.osProgress1.Max = 3
                For i = 0 To 2
                  TotalPasivo = SaldosRazonesCreditos(Fecha2, ListaActivos(i)) + TotalPasivo
                  FrmReportes.osProgress1.Value = FrmReportes.osProgress1.Value + 1
                  DoEvents
                Next
                
                '/////////////////////////TOTAL CAPITAL//////////////////////////////////////////////////////////////////
                TotalCapital = SaldosRazonesCreditos(Fecha2, "Capital")
                
                '////////////////////////TOTAL INGRESOS//////////////////////////////////////////////////////////
                Totalingresos = 0
                Totalingresos = SaldosRazonesCreditos(Fecha2, "Ingresos - Ventas")
                Totalingresos = Totalingresos + SaldosRazonesCreditos(Fecha2, "Servicios - Ventas")
                Totalingresos = Totalingresos + SaldosRazonesCreditos(Fecha2, "Comision - Ventas")
                '//////////////////////////SUMO LOS COSTOS////////////////////////////////////////////////////////
                TotalCosto = 0
                TotalCosto = SaldosRazonesDebitos(Fecha2, "Costos")
'                TotalCosto = TotalCosto + SaldosRazonesDebitos(Fecha2, "Costos Produccion")
'                TotalCosto = TotalCosto + SaldosRazonesDebitos(Fecha2, "Costos Generales Produccion")
                 '//////////////////////////SUMO LOS COSTOS////////////////////////////////////////////////////////
                TotalGastos = 0
                TotalGastos = SaldosRazonesDebitos(Fecha2, "Gastos")
                
                 '//////////////////////////SUMO LAS UTILIDADES////////////////////////////////////////////////////////
                UtilidadBrutas = 0
                UtilidadNetas = 0
                UtilidadBrutas = Totalingresos - TotalCosto
                UtilidadNetas = Totalingresos - TotalCosto - TotalGastos
                
                '/////////////////////TOTAL PASIVO CIRCULANTE//////////////////////////////////////////////////////////////////////////
                i = 0
                TotalPasivoCirculante = 0
                ListaActivos = Array("Proveedores", "Impuestos x Pagar", "Documentos x Pagar CP", "Cobros Anticipados", "Pasivos Acumulados")
                FrmReportes.osProgress1.Value = 0
                FrmReportes.osProgress1.Max = 5
                For i = 0 To 4
                  TotalPasivoCirculante = SaldosRazonesUbicacionCredito(Fecha2, ListaActivos(i)) + TotalPasivoCirculante
                  FrmReportes.osProgress1.Value = FrmReportes.osProgress1.Value + 1
                Next
                
                '/////////////////////TOTAL PASIVO FIJO//////////////////////////////////////////////////////////////////////////
                i = 0
                TotalPasivoFijo = 0
                ListaActivos = Array("Cuentas x Pagar LP", "Documentos x Pagar LP")
                FrmReportes.osProgress1.Value = 0
                FrmReportes.osProgress1.Max = 2
                
                For i = 0 To 1
                  TotalPasivoFijo = SaldosRazonesUbicacionCredito(Fecha2, ListaActivos(i)) + TotalPasivoFijo
                  FrmReportes.osProgress1.Value = FrmReportes.osProgress1.Value + 1
                  DoEvents
                Next
                
                '/////////////////////TOTAL CAPITAL CONTABLE/////////////////////////////////////////////
                 TotalCapitalSocial = SaldosRazonesUbicacionCredito(Fecha2, "Acciones Comunes") + TotalCapitalSocial
               
                 '/////////////////////TOTAL CUENTAS X PAGAR/////////////////////////////////////////////
                 TotalCuentasxPagar = SaldosRazonesCreditos(Fecha2, "Cuentas x Pagar") + TotalCuentasxPagar
         
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



                Me.LblValorCapital.Caption = Format(TotalActivoCirculante, "##,##0.00") & " - " & Format(TotalPasivoCirculante, "##,##0.00")
                Me.LblRazonCapital.Caption = Format(TotalActivoCirculante - TotalPasivoCirculante, "##,##0.00")
                
                Me.LblActivoCirculanteLiquidez.Caption = Format(TotalActivoCirculante, "##,##0.00")
                Me.LblPasivoCirculanteLiquidez.Caption = Format(TotalPasivoCirculante, "##,##0.00")
                If TotalPasivoCirculante = 0 Then
                 Me.LblRazonLiquidez.Caption = 0
                Else
                 Me.LblRazonLiquidez.Caption = Format(TotalActivoCirculante / TotalPasivoCirculante, "##,##0.00")
                End If
                
                Me.LblActivoCirculanteAcido.Caption = Format(TotalActivoCirculante, "##,##0.00") & " - " & Format(TotalInventario, "##,##0.00")
                Me.LblPasivoCirculanteAcido.Caption = Format(TotalPasivoCirculante, "##,##0.00")
                If TotalPasivoCirculante = 0 Then
                 Me.LblRazonAcido.Caption = 0
                Else
                 Me.LblRazonAcido.Caption = Format((TotalActivoCirculante - TotalInventario) / TotalPasivoCirculante, "##,##0.00")
                End If
                
                Me.LblValorCapitalContableOrigen.Caption = Format(TotalCapital, "##,##0.00")
                Me.LblValorPasivoTotalOrigen.Caption = Format(TotalPasivo, "##,##0.00")
                If TotalCapital = 0 Then
                 Me.LblRazonOrigen.Caption = 0
                Else
                 Me.LblRazonOrigen.Caption = Format(TotalPasivo / TotalCapital, "##,##0.00")
                End If
                
                               
                Me.LblValorCapitalContableOrigenCP = Format(TotalCapital, "##,##0.00")
                Me.LblValorPasivoCirculanteOrigen.Caption = Format(TotalPasivoCirculante, "##,##0.00")
                If TotalCapital = 0 Then
                 Me.LblRazonOrigenCapital.Caption = 0
                Else
                 Me.LblRazonOrigenCapital.Caption = Format(TotalPasivoCirculante / TotalCapital, "##,##0.00")
                End If
                
                Me.LblValorPasivoTotalLP.Caption = Format(TotalPasivoFijo, "##,##0.00")
                Me.LblValorCapitalLP.Caption = Format(TotalCapital, "##,##0.00")
                If TotalCapital = 0 Then
                 Me.LblRazonCapitalLP.Caption = 0
                Else
                 Me.LblRazonCapitalLP.Caption = Format(TotalPasivoFijo / TotalCapital, "##,##0.00")
                End If
                
                Me.LblValorActivoFijoInversion.Caption = Format(TotalActivoFijo, "##,##0.00")
                Me.LblValorCapìtalContableInversion.Caption = Format(TotalCapital, "##,##0.00")
                If TotalCapital = 0 Then
                 Me.LblRazonInversion.Caption = 0
                Else
                 Me.LblRazonInversion.Caption = Format(TotalActivoFijo / TotalCapital, "##,##0.00")
                End If
                
                Me.LblValorCostoVentasRotacion.Caption = Format(TotalCosto, "##,##0.00")
                Me.LblValorInventarioRotacion.Caption = Format(TotalInventario, "##,##0.00")
                If TotalInventario = 0 Then
                 Me.LblRazonRotacion.Caption = 0
                 Me.LblRazonIndice.Caption = 0
                Else
                 Me.LblRazonRotacion.Caption = Format(TotalCosto / TotalInventario, "##,##0.00")
                 If TotalCosto <> 0 Then
                  Me.LblRazonIndice.Caption = Format(360 / (TotalCosto / TotalInventario), "##,##0.00")
                 Else
                  MsgBox "El Costo Es Cero, Para las Razones no pueden ser Cero", vbCritical, "Sistema Contable"
                 End If
                 Me.LblValorRotacionIndice.Caption = Format(TotalCosto / TotalInventario, "##,##0.00")
                End If
                
                Me.LblValorCuentasxCobrarCobranza.Caption = Format(TotalCuentaxCobrar, "##,##0.00")
                Me.LblValorVentasCobranza.Caption = Format(Totalingresos, "##,##0.00")
                If Totalingresos = 0 Then
                 Me.LblRazonCobranza.Caption = 0
                 
                Else
                 Me.LblRazonCobranza.Caption = Format(TotalCuentaxCobrar / Totalingresos, "##,##0.00")
                End If
                
                Me.LblValorCuentasxPagarPago.Caption = Format(TotalCuentasxPagar, "##,##0.00")
                Me.LblValorComprasPago.Caption = Format(TotalCompras, "##,##0.00")
                If TotalCompras = 0 Then
                 Me.LblRazonPago.Caption = 0
                 
                Else
                 Me.LblRazonPago.Caption = Format(TotalCuentasxPagar / TotalCompras, "##,##0.00")
                End If
          
                Me.LblValorActivoFijoRotacion.Caption = Format(TotalActivoFijo, "##,##0.00")
                Me.LblValorVentasRotacion.Caption = Format(Totalingresos, "##,##0.00")
                If TotalActivoFijo = 0 Then
                 Me.LblRazonActivoFijo.Caption = 0
                Else
                 Me.LblRazonActivoFijo.Caption = Format(Totalingresos / TotalActivoFijo, "##,##0.00")
                End If
                
                Me.LblValorActivosTotalesAT.Caption = Format(TotalActivos, "##,##0.00")
                Me.LblValorVentasAT.Caption = Format(Totalingresos, "##,##0.00")
                If TotalActivos = 0 Then
                 Me.LblRazonAT.Caption = 0
                Else
                 Me.LblRazonAT.Caption = Format(Totalingresos / TotalActivos, "##,##0.00")
                End If
                
                Me.LblValorActivosTotalesDeuda.Caption = Format(TotalActivos, "##,##0.00")
                Me.LblValorPasivosTotalesDeuda.Caption = Format(TotalPasivo, "##,##0.00")
                If TotalActivos = 0 Then
                 Me.LblRazonDeuda.Caption = 0
                Else
                 Me.LblRazonDeuda.Caption = Format(TotalPasivo / TotalActivos, "##,##0.00")
                End If
                
                Me.LblValorUtilidadBrutaMargen.Caption = Format(UtilidadBrutas, "##,##0.00")
                Me.LblValorVentasMargen.Caption = Format(Totalingresos, "##,##0.00")
                If Totalingresos = 0 Then
                 Me.LblRazonMargen.Caption = 0
                Else
                 Me.LblRazonMargen.Caption = Format(UtilidadBrutas / Totalingresos, "##,##0.00")
                End If
                
                Me.LblValorUtilidadBrutaNeto.Caption = Format(UtilidadNetas, "##,##0.00")
                Me.LblValorVentasNeto.Caption = Format(Totalingresos, "##,##0.00")
                If Totalingresos = 0 Then
                 Me.LblRazonNeto.Caption = 0
                Else
                 Me.LblRazonNeto.Caption = Format(UtilidadNetas / Totalingresos, "##,##0.00")
                End If
                
                
                Me.LblValorUtilidadRI.Caption = Format(UtilidadNetas, "##,##0.00")
                Me.LblValorCapitalSocialRI.Caption = Format(TotalCapitalSocial, "##,##0.00")
                If TotalCapitalSocial = 0 Then
                 Me.LblRazonRI.Caption = 0
                Else
                 Me.LblRazonRI.Caption = Format(UtilidadNetas / TotalCapitalSocial, "##,##0.00")
                End If
                
                Me.LblValorVentasCC.Caption = Format(Totalingresos, "##,##0.00")
                Me.LblValorCapitalContableCC.Caption = Format(TotalCapital, "##,##0.00")
                If TotalCapital = 0 Then
                 Me.LblRazonCC.Caption = 0
                Else
                 Me.LblRazonCC.Caption = Format(Totalingresos / TotalCapital, "##,##0.00")
                End If
                
                 Me.LblValorVentasCN.Caption = Format(Totalingresos, "##,##0.00")
                 Me.LblValorCapitalNetoCN.Caption = Format(TotalCapital + UtilidadNetas, "##,##0.00")
                 If TotalCapital + UtilidadNetas = 0 Then
                   Me.LblRazonCN.Caption = 0
                 Else
                   Me.LblRazonCN.Caption = Format(Totalingresos / (TotalCapital + UtilidadNetas), "##,##0.00")
                 End If
                 
                Me.LblValorUtilidadNetaRCC.Caption = Format(UtilidadNetas, "##,##0.00")
                Me.LblValorCapitalContableRCC.Caption = Format(TotalCapital, "##,##0.00")
                If TotalCapital = 0 Then
                 Me.LblRazonRCC.Caption = 0
                Else
                 Me.LblRazonRCC.Caption = Format(UtilidadNetas / TotalCapital, "##,##0.00")
                End If
                
                Me.LblValorUtilidadNetaRAT.Caption = Format(UtilidadNetas, "##,##0.00")
                Me.LblValorActivoTotalRACT.Caption = Format(TotalActivos, "##,##0.00")
                If TotalActivos = 0 Then
                 Me.LblRazonRACT.Caption = 0
                Else
                 Me.LblRazonRACT.Caption = Format(UtilidadNetas / TotalActivos, "##,##0.00")
                End If
                
                Me.LblValorUtilidadNetaMU.Caption = Format(UtilidadNetas, "##,##0.00")
                Me.LblValorVentasNetasMU.Caption = Format(Totalingresos, "##,##0.00")
                If Totalingresos = 0 Then
                  Me.LblRazonMU.Caption = 0
                Else
                  Me.LblRazonMU.Caption = Format(UtilidadNetas / Totalingresos, "##,##0.00")
                End If
                
                Me.LblMoneda.Caption = "Expresado en " & FrmReportes.CmbMoneda.Text
                Me.Logo.Picture = LoadPicture(RutaLogo)
                Me.LblEmpresa = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
                Me.LblEmpresa1 = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
                Me.LblEmpresa2 = "RUC: " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
                Me.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
                Me.LblFechaFin = FechaFin
                Me.LblFechaIni = FechaIni
End Sub

