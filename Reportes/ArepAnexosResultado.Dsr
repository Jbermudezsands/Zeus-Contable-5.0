VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepAnexosResultado 
   Caption         =   "Reporte de Anexos de Resultado"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   33655
   _ExtentY        =   20267
   SectionData     =   "ArepAnexosResultado.dsx":0000
End
Attribute VB_Name = "ArepAnexosResultado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
  Dim EncabezadoConsulta As String, Condiciones As String
  Dim SQL As String
  
    FrmReportes.DtaReportes.Refresh
    Me.Logo.Picture = LoadPicture(RutaLogo)
    Me.LblEmpresa = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
    Me.LblEmpresa1 = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
    Me.LblEmpresa2 = "RUC: " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
    Me.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
    Me.LblFechaFin = FechaFin
    Me.LblFechaIni = FechaIni
  
  
  '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  '/////////////////////////////GUARDO EL ENZABEZADO DE LA CONSULTA/////////////////////////////////////////////////////
  '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  EncabezadoConsulta = "SELECT Reportes.Descripcion AS Descripcion, Reportes.Debe1 AS Debe1, Reportes.Haber1 AS Haber1, Reportes.Debe2 AS Debe2,Reportes.Haber2 AS Haber2, Reportes.Debe3 AS Debe3, Reportes.Haber3 AS Haber3, Reportes.KeyGrupo AS KeyGrupo,Reportes.KeyGrupoSuperior AS KeyGrupoSuperior, Reportes.KeyGrupoCuenta AS KeyGrupoCuenta, Reportes.Nivel AS Nivel, Reportes.Orden AS Orden,Reportes.CodCuentas AS CodCuentas, Cuentas.DescripcionCuentas, Reportes.Ubicacion, Cuentas.TipoCuenta  " & _
                       "FROM Reportes INNER JOIN Cuentas ON Reportes.KeyGrupo = Cuentas.CodCuentas  "
  
  Condiciones = ""
  
  If FrmReportes.ChkVentas.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Ingresos - Ventas') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Ingresos - Ventas')"
    End If
  End If
  
  If FrmReportes.ChkVentasServicios.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Servicios - Ventas') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Servicios - Ventas')"
    End If
  End If
  
  If FrmReportes.ChkIngresoComisiones.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Comision - Ventas') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Comision - Ventas')"
    End If
  End If
  
  If FrmReportes.ChkRebajas.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Rebajas y Dev S/Venta') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Rebajas y Dev S/Venta')"
    End If
  End If
  
'  If FrmReportes.chkcom = xtpChecked Then
'    If Condiciones = "" Then
'      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Compras') "
'    Else
'      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Compras')"
'    End If
'  End If
  
  If FrmReportes.ChkCostoVentas.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Costos') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Costos')"
    End If
  End If
  
  If FrmReportes.ChkCostoProduccion.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Costos Produccion') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Costos Produccion')"
    End If
  End If
  
  If FrmReportes.ChkCostosGeneralesProduccion.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Costos Generales Produccion') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Costos Generales Produccion')"
    End If
  End If

'  If FrmReportes.chk = xtpChecked Then
'    If Condiciones = "" Then
'      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Acarreo y Fletes') "
'    Else
'      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Acarreo y Fletes')"
'    End If
'  End If
  
'   If FrmReportes.ChkPagosAnticipados.Value = xtpChecked Then
'    If Condiciones = "" Then
'      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Rebajas y Dev S/Compra') "
'    Else
'      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Rebajas y Dev S/Compra')"
'    End If
'   End If
   
   If FrmReportes.ChkComisiones.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Sueldos y Comisiones') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Sueldos y Comisiones')"
    End If
   End If
   
   If FrmReportes.ChkPropaganda.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Propaganda') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Propaganda')"
    End If
   End If

'   If FrmReportes.chkga = xtpChecked Then
'    If Condiciones = "" Then
'      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Gastos') "
'    Else
'      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Gastos')"
'    End If
'   End If
   
   If FrmReportes.ChkSueldosAdmon.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Sueldos Admon') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Sueldos Admon')"
    End If
   End If
   
   If FrmReportes.ChkEnergiaElectrica.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Energia y Agua Potable') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Energia y Agua Potable')"
    End If
   End If
   
   If FrmReportes.ChkComisionesGanadas.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Comisiones/Intereses Gandados') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Comisiones/Intereses Gandados')"
    End If
   End If
   
   If FrmReportes.ChkComisionesPagadas.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Comisiones/Intereses Pagados') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Comisiones/Intereses Pagados')"
    End If
   End If
   
   If FrmReportes.ChkOtrosIngresos.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Otros Ingresos') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Otros Ingresos')"
    End If
   End If
   
   If FrmReportes.ChkOtrosGastos.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Otros Gastos') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Otros Gastos')"
    End If
   End If
   
  
   If FrmReportes.ChkImpuestosxPagar.Value = xtpChecked Then
    If Condiciones = "" Then
      Condiciones = "WHERE (Reportes.Ubicacion IS NOT NULL) AND (Reportes.Ubicacion = N'Impuestos Pagados') "
    Else
      Condiciones = Condiciones & " OR (Reportes.Ubicacion = N'Impuestos Pagados')"
    End If
   End If
   
If Condiciones <> "" Then
 SQL = EncabezadoConsulta & Condiciones & " ORDER BY Reportes.Orden"
    Me.Logo.Picture = LoadPicture(RutaLogo)
'    me.LblMoneda.Caption = "Expresado en " & FrmReportes.CmbMoneda.Text
    Me.LblEmpresa = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
    Me.LblEmpresa1 = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
    Me.LblEmpresa2 = "RUC: " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
    Me.LblFechaFin = Format(FechaFin, "dd/mm/yyyy")
    Me.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
    Me.LblFechaIni = Format(FechaIni, "dd/mm/yyyy")
    Me.DataControl1.ConnectionString = ConexionReporte
    Me.DataControl1.Source = SQL

End If
End Sub

Private Sub Detail_Format()
Dim TipoCuenta As String, Debito As Double, Credito As Double

TipoCuenta = Me.FldTipoCuenta.Text
If Me.FldDebito.Text = "" Then
 Debito = 0
Else
 Debito = Me.FldDebito.Text
End If

If Me.FldCredito.Text = "" Then
 Credito = 0
Else
 Credito = Me.FldCredito.Text
End If

If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
   Me.LblSaldo.Caption = Format(Debito - Credito, "##,##0.00")
Else
   Me.LblSaldo.Caption = Format(Debito - Credito, "##,##0.00")
End If
End Sub

Private Sub GroupFooter1_Format()
Dim TipoCuenta As String, Debito As Double, Credito As Double

TipoCuenta = Me.FldUbicacion.Text
If Me.FldDebitoTotal.Text = "" Then
 Debito = 0
Else
 Debito = Me.FldDebitoTotal.Text
End If

If Me.FldCreditoTotal.Text = "" Then
 Credito = 0
Else
 Credito = Me.FldCreditoTotal.Text
End If

If TipoCuenta = "Cajas" Or TipoCuenta = "Bancos" Or TipoCuenta = "Inventario" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Terreno y Edificios" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Mobiliario y Equipo de Oficina" Or TipoCuenta = "Equipo Rodante" Or TipoCuenta = "Depreciacion Acumulada" Or TipoCuenta = "Papeleria y Utiles de Oficina" Or TipoCuenta = "Pagos Anticipados" Or TipoCuenta = "Otros Activos" Then
   Me.LblSaldoTotal.Caption = Format(Debito - Credito, "##,##0.00")
Else
   Me.LblSaldoTotal.Caption = Format(Debito - Credito, "##,##0.00")
End If
End Sub

