VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepBalancePersonalizado 
   Caption         =   "Balance General - Resumen"
   ClientHeight    =   11490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   19080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   33655
   _ExtentY        =   20267
   SectionData     =   "ArepBalancePersonalizado.dsx":0000
End
Attribute VB_Name = "ArepBalancePersonalizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ActiveReport_ReportEnd()
 On Error GoTo TipoErrs

If FrmReportes.ChkExportar.Value = 1 Then
Dim myExportObject As ActiveReportsExcelExport.ARExportExcel
Dim Nombre As String
    
Set myExportObject = CreateObject("ActiveReportsExcelExport.ARExportExcel")
myExportObject.FileName = RutaArchivo
'myExportObject.FileName = FrmReportes.CommonDialog1.FileName + ".xls"
myExportObject.Export Me.Pages
Set myExportObject = Nothing

MsgBox "Se ha Exportado con Exito!!!!"

End If

Exit Sub
TipoErrs:
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
Dim mes As Double, TipoReporte As String


    TipoReporte = FrmReportes.TxtTipoReporte.Text
    Me.Logo.Picture = LoadPicture(RutaLogo)
    Me.LblMoneda.Caption = "Expresado en " & FrmReportes.CmbMoneda.Text
    Me.LblEmpresa = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
    Me.LblEmpresa1 = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
    Me.LblEmpresa2 = "RUC: " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
    Me.LblFechaFin = Format(FechaFin, "dd/mm/yyyy")
    Me.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
    Me.LblFechaIni = Format(FechaIni, "dd/mm/yyyy")

    '////////////////////RESULTADOS DE ACTIVO CIRCULANTE//////////////////////////////
    TotalActivoCirculante = 0
    SaldosPersonalizados ("Cajas")
    Me.LblCajas.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoCirculante = TotalActivoCirculante + ResultadoPersonalizado
    SaldosPersonalizados ("Bancos")
    Me.LblBancos.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoCirculante = TotalActivoCirculante + ResultadoPersonalizado
    SaldosPersonalizados ("Inventario")
    Me.LblInventario.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoCirculante = TotalActivoCirculante + ResultadoPersonalizado
    SaldosPersonalizados ("Cuentas x Cobrar")
    Me.LblCtasCobrar.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoCirculante = TotalActivoCirculante + ResultadoPersonalizado
    Me.LblTotalActivoCirculante.Caption = Format(TotalActivoCirculante, "##,##0.00")
    
    
    '//////////////////RESULTADOS DE ACTIVO FIJO//////////////////////////////////////////
    SaldosPersonalizados ("Terreno y Edificios")
    Me.LblTerreno.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoFijo = TotalActivoFijo + ResultadoPersonalizado
    SaldosPersonalizados ("Mobiliario y Equipo de Oficina")
    Me.LblMobiliario.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoFijo = TotalActivoFijo + ResultadoPersonalizado
    SaldosPersonalizados ("Equipo Rodante")
    Me.LblEquipoRodante.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoFijo = TotalActivoFijo + ResultadoPersonalizado
    SaldosPersonalizados ("Depreciacion Acumulada")
    Me.LblDepreciacion.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoFijo = TotalActivoFijo + ResultadoPersonalizado
    Me.LblTotalActivoFijo.Caption = Format(TotalActivoFijo, "##,##0.00")
    
    '//////////////////RESULTADOS DE ACTIVO DIFERIDO//////////////////////////////////////////
    SaldosPersonalizados ("Papeleria y Utiles de Oficina")
    Me.LblPapeleria.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoDiferido = TotalActivoDiferido + ResultadoPersonalizado
    SaldosPersonalizados ("Pagos Anticipados")
    Me.LblAnticipos.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoDiferido = TotalActivoDiferido + ResultadoPersonalizado
    SaldosPersonalizados ("Otros Activos")
    Me.LblOtrosActivos.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalActivoDiferido = TotalActivoDiferido + ResultadoPersonalizado
    Me.LblTotalActivoDiferido.Caption = Format(TotalActivoDiferido, "##,##0.00")
    
    '////////////////PASIVO CIRCULANTE//////////////////////////////////////////////////////
    SaldosPersonalizados ("Proveedores")
    Me.LblProveedores.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalPasivoCirculante = TotalPasivoCirculante + ResultadoPersonalizado
    SaldosPersonalizados ("Impuestos x Pagar")
    Me.LblImpuestosPagar.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalPasivoCirculante = TotalPasivoCirculante + ResultadoPersonalizado
    SaldosPersonalizados ("Documentos x Pagar CP")
    Me.LblDocumentosPagar.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalPasivoCirculante = TotalPasivoCirculante + ResultadoPersonalizado
    SaldosPersonalizados ("Cobros Anticipados")
    Me.LblCobrosAnticipados.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalPasivoCirculante = TotalPasivoCirculante + ResultadoPersonalizado
    SaldosPersonalizados ("Pasivos Acumulados")
    Me.LblPasivosAcumulados.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalPasivoCirculante = TotalPasivoCirculante + ResultadoPersonalizado
    Me.LblTotalPasivoCirculante.Caption = Format(TotalPasivoCirculante, "##,##0.00")
    
    '////////////////PASIVO FIJO//////////////////////////////////////////////////////
    SaldosPersonalizados ("Cuentas x Pagar LP")
    Me.LblCuentasPagarLP.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalPasivoFijo = TotalPasivoFijo + ResultadoPersonalizado
    SaldosPersonalizados ("Documentos x Pagar LP")
    Me.LblDocumentosPagarLP.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalPasivoFijo = TotalPasivoFijo + ResultadoPersonalizado
    Me.LblTotalPasivoFijo.Caption = Format(TotalPasivoFijo, "##,##0.00")
    
    '//////////////PASIVO DIFERIDO///////////////////////////////////////////////
    SaldosPersonalizados ("Otros Pasivos")
    Me.LblOtrosPasivos.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalPasivoDiferido = TotalPasivoDiferido + ResultadoPersonalizado
    Me.LblTotalPasivoDiferido.Caption = Format(TotalPasivoDiferido, "##,##0.00")
    
     '//////////////CAPITAL///////////////////////////////////////////////
    SaldosPersonalizados ("Acciones Comunes")
    Me.LblAccionesComunes.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalCapitalSocial = TotalCapitalSocial + ResultadoPersonalizado
    SaldosPersonalizados ("Utilidades Acumuladas")
    Me.LblUtilidadesAcumuladas.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalCapitalSocial = TotalCapitalSocial + ResultadoPersonalizado
    SaldosPersonalizados ("Otras Ctas de Capital")
    Me.LblOtrasCuentasCapital.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalCapitalSocial = TotalCapitalSocial + ResultadoPersonalizado
    SaldosPersonalizados ("Resultado Periodo")
    Me.LblResultadoPeriodo.Caption = Format(ResultadoPersonalizado, "##,##0.00")
    TotalCapitalSocial = TotalCapitalSocial + ResultadoPersonalizado
    Me.LblTotalCapital.Caption = Format(TotalCapitalSocial, "##,##0.00")
    
    Me.LblTotalPasivomasCapital.Caption = Format(TotalPasivoCirculante + TotalPasivoFijo + TotalPasivoDiferido + TotalCapitalSocial, "##,##0.00")
    Me.LblTotalActivo.Caption = Format(TotalActivoCirculante + TotalActivoFijo + TotalActivoDiferido, "##,##0.00")
    
    If TipoReporte = "BALANCE GENERAL RESUMEN ANEXOS" Then
        FrmReportes.DtaConsulta.RecordSource = "SELECT * From ConfiguracionReporte"
        FrmReportes.DtaConsulta.Refresh
        If Not FrmReportes.DtaConsulta.Recordset.EOF Then
         Me.LblNombreCaja.Caption = FrmReportes.DtaConsulta.Recordset("Caja")
         Me.LblNombreBanco.Caption = FrmReportes.DtaConsulta.Recordset("Banco")
         Me.LblNombreCuentasxCobrar.Caption = FrmReportes.DtaConsulta.Recordset("CtasxCobrar")
         Me.LblNombreInventario.Caption = FrmReportes.DtaConsulta.Recordset("Inventario")
         Me.LblNombreTerrenos.Caption = FrmReportes.DtaConsulta.Recordset("Terreno")
         Me.LblNombreMobiliarios.Caption = FrmReportes.DtaConsulta.Recordset("Mobiliario")
         Me.LblNombreEquipoRodante.Caption = FrmReportes.DtaConsulta.Recordset("EquipoRodante")
         Me.LblNombreDepreciacion.Caption = FrmReportes.DtaConsulta.Recordset("DepAcumulada")
         Me.LblNombrePapeleria.Caption = FrmReportes.DtaConsulta.Recordset("Papeleria")
         Me.LblNombrePagosAnticipados.Caption = FrmReportes.DtaConsulta.Recordset("PagosAnticipados")
         Me.LblNombreOtrosActivos.Caption = FrmReportes.DtaConsulta.Recordset("OtrosActivos")
         Me.LblNombreProveedores.Caption = FrmReportes.DtaConsulta.Recordset("Proveedores")
         Me.LblNombreImpuestosxPagar.Caption = FrmReportes.DtaConsulta.Recordset("ImpuestosxPagar")
         Me.LblNombreDocumentosxPagar.Caption = FrmReportes.DtaConsulta.Recordset("DocumentosxPagar")
         Me.LblNombreCobrosAnticipados.Caption = FrmReportes.DtaConsulta.Recordset("CobroAnticipados")
         Me.LblNombrePasivosAcumulados.Caption = FrmReportes.DtaConsulta.Recordset("PasivosAcumulados")
         Me.LblNombreCuentasLP.Caption = FrmReportes.DtaConsulta.Recordset("PagosLP")
         Me.LblNombreDocumentosLP.Caption = FrmReportes.DtaConsulta.Recordset("DocumentosLP")
         Me.LblNombreOtrosPasivos.Caption = FrmReportes.DtaConsulta.Recordset("OtrosPasivos")
         Me.LblNombreAccionesComunes.Caption = FrmReportes.DtaConsulta.Recordset("AccionesComunes")
         Me.LblNombreUtilidades.Caption = FrmReportes.DtaConsulta.Recordset("UtilidadAcumulada")
         Me.LblNombreOtrasCtasCapital.Caption = FrmReportes.DtaConsulta.Recordset("OtrosCapitales")
    
        End If
    End If

End Sub

