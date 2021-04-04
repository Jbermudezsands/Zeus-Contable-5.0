VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepAnexosBalances 
   Caption         =   "Anexos del Balance General"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   33655
   _ExtentY        =   20267
   SectionData     =   "ArepAnexosBalances.dsx":0000
End
Attribute VB_Name = "ArepAnexosBalances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ActiveReport_ReportStart()
  Dim EncabezadoConsulta As String, Condiciones As String
  Dim SQL As String
  
  '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  '/////////////////////////////GUARDO EL ENZABEZADO DE LA CONSULTA/////////////////////////////////////////////////////
  '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  EncabezadoConsulta = "SELECT Reportes.Descripcion AS Descripcion, Reportes.Debe1 AS Debe1, Reportes.Haber1 AS Haber1, Reportes.Debe2 AS Debe2,Reportes.Haber2 AS Haber2, Reportes.Debe3 AS Debe3, Reportes.Haber3 AS Haber3, Reportes.KeyGrupo AS KeyGrupo,Reportes.KeyGrupoSuperior AS KeyGrupoSuperior, Reportes.KeyGrupoCuenta AS KeyGrupoCuenta, Reportes.Nivel AS Nivel, Reportes.Orden AS Orden,Reportes.CodCuentas AS CodCuentas, Cuentas.DescripcionCuentas, Reportes.Ubicacion, Cuentas.TipoCuenta  " & _
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
   Me.LblSaldo.Caption = Format(Credito - Debito, "##,##0.00")
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
   Me.LblSaldoTotal.Caption = Format(Credito - Debito, "##,##0.00")
End If


End Sub

