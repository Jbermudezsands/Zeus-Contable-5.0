VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepBalanceTradicional 
   Caption         =   "Balance General Tradicional"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepBalanceTradicional.dsx":0000
End
Attribute VB_Name = "ArepBalanceTradicional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalActivoCirculante As Double, TotalActivoFijo As Double, TotalActivoDiferido As Double
Public TotalPasivo As Double, TotalPasivoCirculante As Double, TotalPasivoFijo As Double, TotalPasivoDiferido As Double, TotalCapitalSocial As Double, TotalCapital As Double, Utilidad As Double




Private Sub ActiveReport_ReportEnd()
 On Error GoTo err:

If FrmReportes.ChkExportar.Value = 1 Then
  
    MousePointer = 11
    
    Dim myExportObject As ActiveReportsExcelExport.ARExportExcel
    Dim Nombre As String
    
    Set myExportObject = CreateObject("ActiveReportsExcelExport.ARExportExcel")
    myExportObject.FileName = FrmReportes.CommonDialog1.FileName + ".xls"
    myExportObject.Export Me.Pages
    Set myExportObject = Nothing
    MsgBox "Exportacion con Exito!!!"
End If
err:
    If err.Number <> 0 Then Exit Sub
End Sub

Private Sub ActiveReport_ReportStart()
                
                Me.LblMoneda.Caption = "Expresado en " & FrmReportes.CmbMoneda.Text
                Me.Logo.Picture = LoadPicture(RutaLogo)
                Me.LblEmpresa = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
                Me.LblEmpresa1 = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
                Me.LblEmpresa2 = "RUC: " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
                Me.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
                Me.LblFecha2.Caption = Format(FechaFin, "yyyy-mm-dd")
                Me.LblFechaFin = FechaFin
                Me.LblFechaIni = FechaIni
                
                Me.LblActivoCirculante.Caption = Format(RTotalActivoCirculante, "##,##0.00")
                Me.LblTotalActivoFijo.Caption = Format(RTotalActivoFijo, "##,##0.00")
                Me.LblTotalActivos.Caption = Format(RTotalActivoCirculante + RTotalActivoFijo, "##,##0.00")
                Me.LblTotalPasivo.Caption = Format(RTotalPasivo, "##,##0.00")
                Me.LblTotalCapital.Caption = Format(RTotalCapital, "##,##0.00")
                Me.LblTotalPasivomasCapital.Caption = Format(RTotalCapital + RTotalPasivo + RUtilidad, "##,##0.00")
                Me.LblResultadoPeriodo.Caption = Format(RUtilidad, "##,##0.00")
                
End Sub

Private Sub Detail_Format()
Dim Descripcion As String, SqlActivos As String, KeyGrupo As String, Fecha2 As String
Dim TotalActivoCirculante As Double
FrmReportes.DtaConsulta.RecordSource = "SELECT  * From Reportes Where (Nivel = 2) And (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 = 0) ORDER BY Orden"
FrmReportes.DtaConsulta.Refresh

Fecha2 = Format(FechaFin, "yyyy-mm-dd") 'Me.LblFecha2.Caption
 

 '////////////////////////////////////BUSCO DE LAS CUENTAS DE ACTIVOS////////////////////////////////////////////
    KeyGrupo = "A0100"
'    SqlActivos = "SELECT  Transacciones.CodCuentas AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Cuentas.DescripcionCuentas) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito,SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Debito) - SUM(Transacciones.TCambio * Transacciones.Credito) AS Saldo, Cuentas.KeyGrupo AS KeyGrupo, Cuentas.DescripcionGrupo AS DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta " & _
'                 "FROM  Grupos INNER JOIN Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo  " & _
'                 "WHERE   (Transacciones.FechaTransaccion <= CONVERT(DATETIME,  '" & Fecha2 & "', 102)) GROUP BY Cuentas.TipoCuenta, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, Transacciones.CodCuentas " & _
'                 "HAVING  (Cuentas.TipoCuenta = 'Caja') AND (SUM(Transacciones.TCambio * Transacciones.Debito) - SUM(Transacciones.TCambio * Transacciones.Credito) <> 0) OR (Cuentas.TipoCuenta = 'Bancos') OR (Cuentas.TipoCuenta = 'Cuentas x Cobrar') OR (Cuentas.TipoCuenta = 'Inventario') OR (Cuentas.TipoCuenta = 'Otros Activos') ORDER BY Cuentas.KeyGrupo, Transacciones.CodCuentas"
                 
    SqlActivos = "SELECT   Transacciones.CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Cuentas.DescripcionCuentas) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Debito) - SUM(Transacciones.TCambio * Transacciones.Credito) AS Saldo, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta " & _
                 "FROM  Grupos INNER JOIN Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo " & _
                 "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Cuentas.TipoCuenta, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, Transacciones.CodCuentas HAVING (Cuentas.TipoCuenta = 'Caja') AND (SUM(Transacciones.TCambio * Transacciones.Debito) - SUM(Transacciones.TCambio * Transacciones.Credito) <> 0) OR (Cuentas.TipoCuenta = 'Bancos') AND (SUM(Transacciones.TCambio * Transacciones.Debito) - SUM(Transacciones.TCambio * Transacciones.Credito) <> 0) OR (Cuentas.TipoCuenta = 'Cuentas x Cobrar') AND (SUM(Transacciones.TCambio * Transacciones.Debito) - SUM(Transacciones.TCambio * Transacciones.Credito) <> 0) OR (Cuentas.TipoCuenta = 'Inventario') AND (SUM(Transacciones.TCambio * Transacciones.Debito) - SUM(Transacciones.TCambio * Transacciones.Credito) <> 0) OR (Cuentas.TipoCuenta = 'Otros Activos') AND (SUM(Transacciones.TCambio * Transacciones.Debito) - SUM(Transacciones.TCambio * Transacciones.Credito) <> 0) ORDER BY Cuentas.KeyGrupo, Transacciones.CodCuentas"
                                            
    Set Me.SrptActivoCirculante.object = New ArepBalanceDetalle
    Me.SrptActivoCirculante.object.LblTotalGrupo.Caption = "Total " '& Descripcion
    Me.SrptActivoCirculante.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptActivoCirculante.object.DataControl1.Source = SqlActivos
    
   
    
   
 '////////////////////////////////////BUSCO DE LAS CUENTAS DE ACTIVO FIJOS////////////////////////////////////////////
    KeyGrupo = "A0100"
'    SqlActivos = "SELECT Transacciones.CodCuentas AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta)  AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Debito)-SUM(Transacciones.TCambio * Transacciones.Credito) AS Saldo, Cuentas.KeyGrupo AS KeyGrupo, Cuentas.DescripcionGrupo AS DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta " & _
'                 "FROM Grupos INNER JOIN Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo " & _
'                 "WHERE  (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Cuentas.TipoCuenta, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo,Transacciones.CodCuentas " & _
'                 "HAVING (Cuentas.TipoCuenta = 'Activo Fijo') ORDER BY Cuentas.KeyGrupo"

    SqlActivos = "SELECT  Transacciones.CodCuentas AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Cuentas.DescripcionCuentas) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito,SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Debito) - SUM(Transacciones.TCambio * Transacciones.Credito) AS Saldo, Cuentas.KeyGrupo AS KeyGrupo, Cuentas.DescripcionGrupo AS DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta " & _
                 "FROM  Grupos INNER JOIN Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo  " & _
                 "WHERE   (Transacciones.FechaTransaccion <= CONVERT(DATETIME,  '" & Fecha2 & "', 102)) GROUP BY Cuentas.TipoCuenta, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, Transacciones.CodCuentas " & _
                 "HAVING   (Cuentas.TipoCuenta = 'Activo Fijo') AND (SUM(Transacciones.TCambio * Transacciones.Debito) - SUM(Transacciones.TCambio * Transacciones.Credito) <> 0) ORDER BY Cuentas.KeyGrupo,Transacciones.CodCuentas "
                 
                                            
    Set Me.SrptActivoFijo.object = New ArepBalanceDetalle
    Me.SrptActivoFijo.object.LblTotalGrupo.Caption = "Total " '& Descripcion
    Me.SrptActivoFijo.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptActivoFijo.object.DataControl1.Source = SqlActivos
    
    
 '////////////////////////////////////////BUSCO LAS CUENTAS DE PASIVO///////////////////////////////////////////////////
    KeyGrupo = "B"
                                            
'     SqlActivos = "SELECT  Transacciones.CodCuentas AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Cuentas.DescripcionCuentas) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Credito) - SUM(Transacciones.TCambio * Transacciones.Debito) AS Saldo, Cuentas.KeyGrupo AS KeyGrupo, Cuentas.DescripcionGrupo AS DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta  " & _
'                  "FROM Grupos INNER JOIN  Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo  " & _
'                  "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Cuentas.TipoCuenta, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, Transacciones.CodCuentas " & _
'                  "HAVING (Cuentas.TipoCuenta = 'Cuentas x Pagar') AND (SUM(Transacciones.TCambio * Transacciones.Debito) - SUM(Transacciones.TCambio * Transacciones.Credito) <> 0) OR (Cuentas.TipoCuenta = 'Otros Pasivos') OR (Cuentas.TipoCuenta = 'Pasivo') ORDER BY Cuentas.KeyGrupo,Transacciones.CodCuentas "
      SqlActivos = "SELECT Transacciones.CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Cuentas.DescripcionCuentas) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Credito) - SUM(Transacciones.TCambio * Transacciones.Debito) AS Saldo, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta FROM Grupos INNER JOIN Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo  " & _
                   "WHERE  (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Cuentas.TipoCuenta, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, Transacciones.CodCuentas HAVING (Cuentas.TipoCuenta = 'Cuentas x Pagar') AND (SUM(Transacciones.TCambio * Transacciones.Debito) - SUM(Transacciones.TCambio * Transacciones.Credito) <> 0) OR (Cuentas.TipoCuenta = 'Otros Pasivos') AND (SUM(Transacciones.TCambio * Transacciones.Debito) - SUM(Transacciones.TCambio * Transacciones.Credito) <> 0) OR (Cuentas.TipoCuenta = 'Pasivo') AND (SUM(Transacciones.TCambio * Transacciones.Debito) - SUM(Transacciones.TCambio * Transacciones.Credito) <> 0) ORDER BY Cuentas.KeyGrupo, Transacciones.CodCuentas"

                                            
    Set Me.SrptPasivo.object = New ArepBalanceDetalle
    Me.SrptPasivo.object.LblTotalGrupo.Caption = "Total " '& Descripcion
    Me.SrptPasivo.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptPasivo.object.DataControl1.Source = SqlActivos
 
  '////////////////////////////////////////BUSCO LAS CUENTAS DE CAPITAL///////////////////////////////////////////////////
    KeyGrupo = "C"
'    SqlActivos = "SELECT Transacciones.CodCuentas AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta)  AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Credito) - SUM(Transacciones.TCambio * Transacciones.Debito) AS Saldo, Cuentas.KeyGrupo AS KeyGrupo, Cuentas.DescripcionGrupo AS DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta " & _
'                 "FROM Grupos INNER JOIN Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo " & _
'                 "WHERE  (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Cuentas.TipoCuenta, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo,Transacciones.CodCuentas " & _
'                 "HAVING (Cuentas.TipoCuenta = 'Capital') ORDER BY Cuentas.KeyGrupo"

     SqlActivos = "SELECT  Transacciones.CodCuentas AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Cuentas.DescripcionCuentas) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Credito) - SUM(Transacciones.TCambio * Transacciones.Debito) AS Saldo, Cuentas.KeyGrupo AS KeyGrupo, Cuentas.DescripcionGrupo AS DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta  " & _
                  "FROM Grupos INNER JOIN  Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo  " & _
                  "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Cuentas.TipoCuenta, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, Transacciones.CodCuentas " & _
                  "HAVING (Cuentas.TipoCuenta = 'Capital') AND (SUM(Transacciones.TCambio * Transacciones.Debito) - SUM(Transacciones.TCambio * Transacciones.Credito) <> 0) ORDER BY Cuentas.KeyGrupo,Transacciones.CodCuentas "
                                            
                                            
    Set Me.SrptCapital.object = New ArepBalanceDetalle
    Me.SrptCapital.object.LblTotalGrupo.Caption = "Total " '& Descripcion
    Me.SrptCapital.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptCapital.object.DataControl1.Source = SqlActivos
    
 
 


End Sub

