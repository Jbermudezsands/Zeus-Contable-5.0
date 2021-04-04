VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepCatalogoResumen 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   33655
   _ExtentY        =   20267
   SectionData     =   "ArepCatalgoResumen.dsx":0000
End
Attribute VB_Name = "ArepCatalogoResumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
                Me.LblMoneda.Caption = "Expresado en " & FrmReportes.CmbMoneda.Text
                Me.Logo.Picture = LoadPicture(RutaLogo)
                Me.LblEmpresa = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
                Me.LblEmpresa1 = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
                Me.LblEmpresa2 = "RUC: " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
                Me.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub Detail_Format()
 Dim SqlActivos As String
 
  SqlActivos = "SELECT Cuentas.CodCuentas AS CodCuentas, Cuentas.DescripcionCuentas AS Descripcion, Cuentas.TipoCuenta AS TipoCuenta, Cuentas.TipoMoneda AS TipoMoneda, Grupos.DescripcionGrupo, Grupos.KeyGrupo  " & _
               "FROM  Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo  " & _
               "WHERE (Cuentas.TipoCuenta = N'Caja') OR (Cuentas.TipoCuenta = N'Bancos') OR (Cuentas.TipoCuenta = N'Cuentas x Cobrar') OR (Cuentas.TipoCuenta = N'Inventario') OR  (Cuentas.TipoCuenta = N'Otros Activos') OR (Cuentas.TipoCuenta = N'Papeleria - Utiles') " & _
               "ORDER BY Cuentas.CodCuentas "

   Set Me.SrptActivoCirculante.object = New ArepCatalogoDetalle
    Me.SrptActivoCirculante.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptActivoCirculante.object.DataControl1.Source = SqlActivos
    
   SqlActivos = "SELECT Cuentas.CodCuentas AS CodCuentas, Cuentas.DescripcionCuentas AS Descripcion, Cuentas.TipoCuenta AS TipoCuenta, Cuentas.TipoMoneda AS TipoMoneda, Grupos.DescripcionGrupo, Grupos.KeyGrupo  " & _
               "FROM  Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo  " & _
               "WHERE (Cuentas.TipoCuenta = 'Activo Fijo') " & _
               "ORDER BY Cuentas.CodCuentas "

   Set Me.SrptActivoFijo.object = New ArepCatalogoDetalle
    Me.SrptActivoFijo.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptActivoFijo.object.DataControl1.Source = SqlActivos
    
   SqlActivos = "SELECT Cuentas.CodCuentas AS CodCuentas, Cuentas.DescripcionCuentas AS Descripcion, Cuentas.TipoCuenta AS TipoCuenta, Cuentas.TipoMoneda AS TipoMoneda, Grupos.DescripcionGrupo, Grupos.KeyGrupo  " & _
               "FROM  Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo  " & _
               "WHERE (Cuentas.TipoCuenta = N'Pasivo') OR (Cuentas.TipoCuenta = N'Cuentas x Pagar') OR (Cuentas.TipoCuenta = N'Otros Pasivos') " & _
               "ORDER BY Cuentas.CodCuentas "

   Set Me.SrptPasivo.object = New ArepCatalogoDetalle
    Me.SrptPasivo.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptPasivo.object.DataControl1.Source = SqlActivos
    
   SqlActivos = "SELECT Cuentas.CodCuentas AS CodCuentas, Cuentas.DescripcionCuentas AS Descripcion, Cuentas.TipoCuenta AS TipoCuenta, Cuentas.TipoMoneda AS TipoMoneda, Grupos.DescripcionGrupo, Grupos.KeyGrupo  " & _
               "FROM  Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo  " & _
               "WHERE (Cuentas.TipoCuenta = N'Capital') " & _
               "ORDER BY Cuentas.CodCuentas "

   Set Me.SrptCapital.object = New ArepCatalogoDetalle
    Me.SrptCapital.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptCapital.object.DataControl1.Source = SqlActivos
    
   SqlActivos = "SELECT Cuentas.CodCuentas AS CodCuentas, Cuentas.DescripcionCuentas AS Descripcion, Cuentas.TipoCuenta AS TipoCuenta, Cuentas.TipoMoneda AS TipoMoneda, Grupos.DescripcionGrupo, Grupos.KeyGrupo  " & _
               "FROM  Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo  " & _
               "WHERE (Cuentas.TipoCuenta = N'Ingresos - Ventas') " & _
               "ORDER BY Cuentas.CodCuentas "

   Set Me.SrptIngresos.object = New ArepCatalogoDetalle
    Me.SrptIngresos.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptIngresos.object.DataControl1.Source = SqlActivos
    
   SqlActivos = "SELECT Cuentas.CodCuentas AS CodCuentas, Cuentas.DescripcionCuentas AS Descripcion, Cuentas.TipoCuenta AS TipoCuenta, Cuentas.TipoMoneda AS TipoMoneda, Grupos.DescripcionGrupo, Grupos.KeyGrupo  " & _
               "FROM  Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo  " & _
               "WHERE (Cuentas.TipoCuenta = N'Costos') " & _
               "ORDER BY Cuentas.CodCuentas "

   Set Me.SrptCostos.object = New ArepCatalogoDetalle
    Me.SrptCostos.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptCostos.object.DataControl1.Source = SqlActivos

   SqlActivos = "SELECT Cuentas.CodCuentas AS CodCuentas, Cuentas.DescripcionCuentas AS Descripcion, Cuentas.TipoCuenta AS TipoCuenta, Cuentas.TipoMoneda AS TipoMoneda, Grupos.DescripcionGrupo, Grupos.KeyGrupo  " & _
               "FROM  Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo  " & _
               "WHERE (Cuentas.TipoCuenta = N'Gastos') " & _
               "ORDER BY Cuentas.CodCuentas "

   Set Me.SrptGastos.object = New ArepCatalogoDetalle
    Me.SrptGastos.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptGastos.object.DataControl1.Source = SqlActivos
    
   SqlActivos = "SELECT Cuentas.CodCuentas AS CodCuentas, Cuentas.DescripcionCuentas AS Descripcion, Cuentas.TipoCuenta AS TipoCuenta, Cuentas.TipoMoneda AS TipoMoneda, Grupos.DescripcionGrupo, Grupos.KeyGrupo  " & _
               "FROM  Cuentas INNER JOIN  Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo  " & _
               "WHERE (Cuentas.TipoCuenta = N'Cuentas de Orden') " & _
               "ORDER BY Cuentas.CodCuentas "

   Set Me.SrptCuentasOrden.object = New ArepCatalogoDetalle
    Me.SrptCuentasOrden.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptCuentasOrden.object.DataControl1.Source = SqlActivos



End Sub
