VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepResultadoTradicional 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepResultadoTradicional.dsx":0000
End
Attribute VB_Name = "ArepResultadoTradicional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ActiveReport_ReportStart()
Dim NumeroPeriodo1 As Double, NumeroPeriodo2 As Double, NumeroTabla As Double

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
           FrmReportes.DtaConsulta.Recordset.MoveLast
           i = FrmReportes.DtaConsulta.Recordset.RecordCount
           FrmReportes.DtaConsulta.Recordset.MoveFirst
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
            FrmReportes.DtaConsulta.Recordset.MoveNext
          Loop
        
        FrmReportes.DtaReportes.Refresh
        
        Fecha1 = Format(FechaIni, "yyyy-mm-dd")
        Fecha2 = Format(FechaFin, "yyyy-mm-dd")
        
        Me.LblFecha1.Caption = Fecha1
        Me.LblFecha2.Caption = Fecha2
        
        Me.LblMoneda.Caption = "Expresado en " & FrmReportes.CmbMoneda.Text
        Me.Logo.Picture = LoadPicture(RutaLogo)
        Me.LblEmpresa = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
        Me.LblEmpresa1 = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
        Me.LblEmpresa2 = "RUC: " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
        Me.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
        Me.LblFechaFin = FechaFin
        Me.LblFechaIni = FechaIni
End Sub

Private Sub Detail_Format()
Dim SqlIngresos As String, Fecha1 As String, SqlCostos As String, SqlGastos As String, Fecha2 As String
Dim Totalingresos As Double, TotalCostos As Double, TotalGastos As Double, UtilidadNeta As Double
Fecha1 = Format(Me.LblFecha1.Caption, "yyyy-mm-dd")
Fecha2 = Format(Me.LblFecha2.Caption, "yyyy-mm-dd")


If FrmReportes.CmbMoneda.Text = "Córdobas" Then
    If FrmReportes.OptAcumulado.Value = True Then
      SqlIngresos = "SELECT Transacciones.CodCuentas AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Credito)- SUM(Transacciones.TCambio * Transacciones.Debito) AS Saldo, Cuentas.KeyGrupo AS KeyGrupo, Cuentas.DescripcionGrupo AS DescripcionGrupo,MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta " & _
         "FROM  Grupos INNER JOIN Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo " & _
         "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
         "GROUP BY Cuentas.TipoCuenta, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo,Transacciones.CodCuentas " & _
         "HAVING  (Cuentas.TipoCuenta = 'Ingresos - Ventas') AND (SUM(Transacciones.TCambio * Transacciones.Credito) - SUM(Transacciones.TCambio * Transacciones.Debito) <> 0) ORDER BY Cuentas.KeyGrupo "
    ElseIf FrmReportes.OptPeriodo.Value = True Then
      SqlIngresos = "SELECT Transacciones.CodCuentas AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Credito)- SUM(Transacciones.TCambio * Transacciones.Debito) AS Saldo, Cuentas.KeyGrupo AS KeyGrupo, Cuentas.DescripcionGrupo AS DescripcionGrupo,MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta " & _
         "FROM  Grupos INNER JOIN Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo " & _
         "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
         "GROUP BY Cuentas.TipoCuenta, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo,Transacciones.CodCuentas " & _
         "HAVING  (Cuentas.TipoCuenta = 'Ingresos - Ventas') AND (SUM(Transacciones.TCambio * Transacciones.Credito) - SUM(Transacciones.TCambio * Transacciones.Debito) <> 0) ORDER BY Cuentas.KeyGrupo "
    End If
Else
    If FrmReportes.OptAcumulado.Value = True Then
      SqlIngresos = "SELECT Transacciones.CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta,AVG(Transacciones.TCambio) AS TCambio, SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Debito, 2)) AS Debito,SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 2)) AS Credito,SUM (Round(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 2)) - SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Debito, 2)) AS Saldo, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo,MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta, Tasas.MontoCordobas FROM  Grupos INNER JOIN Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
                    "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
                    "GROUP BY Cuentas.TipoCuenta, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, Transacciones.CodCuentas, Tasas.MontoCordobas " & _
                    "HAVING      (Cuentas.TipoCuenta = 'Ingresos - Ventas') AND (SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 2)) - SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Debito, 2)) <> 0) ORDER BY Cuentas.KeyGrupo"
    ElseIf FrmReportes.OptPeriodo.Value = True Then
      SqlIngresos = "SELECT Transacciones.CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta,AVG(Transacciones.TCambio) AS TCambio, SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Debito, 2)) AS Debito,SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 2)) AS Credito,SUM (Round(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 2)) - SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Debito, 2)) AS Saldo, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo,MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta, Tasas.MontoCordobas FROM  Grupos INNER JOIN Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
                    "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
                    "GROUP BY Cuentas.TipoCuenta, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, Transacciones.CodCuentas, Tasas.MontoCordobas " & _
                    "HAVING      (Cuentas.TipoCuenta = 'Ingresos - Ventas') AND (SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 2)) - SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Debito, 2)) <> 0) ORDER BY Cuentas.KeyGrupo"
    End If
End If
      
    Set Me.SrptDetalleIngresos.object = New ArepResultadoDetalle
    Me.SrptDetalleIngresos.object.LblTotalGrupo.Caption = "Total de Ingresos/Ventas"
    Me.SrptDetalleIngresos.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptDetalleIngresos.object.DataControl1.Source = SqlIngresos
    Totalingresos = ResultadoPersonalizado
    
    
 '////////////////////////////////////////lleno los datos de los costos///////////////////////////////////
     If FrmReportes.CmbMoneda.Text = "Córdobas" Then
        If FrmReportes.OptAcumulado.Value = True Then
          SqlCostos = "SELECT Transacciones.CodCuentas AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Debito)- SUM(Transacciones.TCambio * Transacciones.Credito)  AS Saldo, Cuentas.KeyGrupo AS KeyGrupo, Cuentas.DescripcionGrupo AS DescripcionGrupo,MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta " & _
            "FROM  Grupos INNER JOIN Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo " & _
            "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
            "GROUP BY Cuentas.TipoCuenta, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo,Transacciones.CodCuentas " & _
            "HAVING  (Cuentas.TipoCuenta = 'Costos') AND (SUM(Transacciones.TCambio * Transacciones.Credito) - SUM(Transacciones.TCambio * Transacciones.Debito) <> 0) ORDER BY Cuentas.KeyGrupo "
        
        ElseIf FrmReportes.OptPeriodo.Value = True Then
          SqlCostos = "SELECT Transacciones.CodCuentas AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Debito)- SUM(Transacciones.TCambio * Transacciones.Credito)  AS Saldo, Cuentas.KeyGrupo AS KeyGrupo, Cuentas.DescripcionGrupo AS DescripcionGrupo,MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta " & _
            "FROM  Grupos INNER JOIN Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo " & _
            "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
            "GROUP BY Cuentas.TipoCuenta, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo,Transacciones.CodCuentas " & _
            "HAVING  (Cuentas.TipoCuenta = 'Costos') AND (SUM(Transacciones.TCambio * Transacciones.Credito) - SUM(Transacciones.TCambio * Transacciones.Debito) <> 0) ORDER BY Cuentas.KeyGrupo "
        End If
    Else
        If FrmReportes.OptAcumulado.Value = True Then
          SqlCostos = "SELECT Transacciones.CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta,AVG(Transacciones.TCambio) AS TCambio, SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Debito, 2)) AS Debito,SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 2)) AS Credito,SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Debito, 2))-SUM (Round(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 2)) AS Saldo, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo,MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta, Tasas.MontoCordobas FROM  Grupos INNER JOIN Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
                        "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
                        "GROUP BY Cuentas.TipoCuenta, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, Transacciones.CodCuentas, Tasas.MontoCordobas " & _
                        "HAVING  (Cuentas.TipoCuenta = 'Costos') AND (SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 2)) - SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Debito, 2)) <> 0) ORDER BY Cuentas.KeyGrupo"
        ElseIf FrmReportes.OptPeriodo.Value = True Then
          SqlCostos = "SELECT Transacciones.CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta,AVG(Transacciones.TCambio) AS TCambio, SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Debito, 2)) AS Debito,SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 2)) AS Credito,SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Debito, 2))-SUM (Round(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 2)) AS Saldo, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo,MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta, Tasas.MontoCordobas FROM  Grupos INNER JOIN Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
                        "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
                        "GROUP BY Cuentas.TipoCuenta, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, Transacciones.CodCuentas, Tasas.MontoCordobas " & _
                        "HAVING (Cuentas.TipoCuenta = 'Costos') AND (SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 2)) - SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Debito, 2)) <> 0) ORDER BY Cuentas.KeyGrupo"
        End If
    End If
      
      
    Set Me.SrptDetalleCostos.object = New ArepResultadoDetalle
    Me.SrptDetalleCostos.object.LblTotalGrupo.Caption = "Total de Costos"
    Me.SrptDetalleCostos.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptDetalleCostos.object.DataControl1.Source = SqlCostos
'    TotalCostos = Me.SrptDetalleIngresos.object.FldTotal.Text
    
  '////////////////////////////////////////lleno los datos de los Gastos///////////////////////////////////
   If FrmReportes.CmbMoneda.Text = "Córdobas" Then
          If FrmReportes.OptAcumulado.Value = True Then
           SqlGastos = "SELECT Transacciones.CodCuentas AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Debito)- SUM(Transacciones.TCambio * Transacciones.Credito)  AS Saldo, Cuentas.KeyGrupo AS KeyGrupo, Cuentas.DescripcionGrupo AS DescripcionGrupo,MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta " & _
             "FROM  Grupos INNER JOIN Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo " & _
             "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
             "GROUP BY Cuentas.TipoCuenta, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo,Transacciones.CodCuentas " & _
             "HAVING  (Cuentas.TipoCuenta = 'Gastos') AND (SUM(Transacciones.TCambio * Transacciones.Credito) - SUM(Transacciones.TCambio * Transacciones.Debito) <> 0) ORDER BY Cuentas.KeyGrupo "
          Else
          SqlGastos = "SELECT MAX(Cuentas.TipoCuenta) AS TipoCuenta, Transacciones.CodCuentas AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Debito) - SUM(Transacciones.TCambio * Transacciones.Credito) AS Saldo, MAX(Cuentas.KeyGrupo)  AS KeyGrupo, MAX(Cuentas.DescripcionGrupo) AS DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta  " & _
                      "FROM Grupos INNER JOIN Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo " & _
                      "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Transacciones.CodCuentas " & _
                      "HAVING (SUM(Transacciones.TCambio * Transacciones.Credito) - SUM(Transacciones.TCambio * Transacciones.Debito) <> 0) AND (MAX(Cuentas.TipoCuenta) = N'Gastos') ORDER BY MAX(Cuentas.KeyGrupo)"
          End If
   Else
        If FrmReportes.OptAcumulado.Value = True Then
          SqlGastos = "SELECT Transacciones.CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta,AVG(Transacciones.TCambio) AS TCambio, SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Debito, 2)) AS Debito,SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 2)) AS Credito,SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Debito, 2))-SUM (Round(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 2)) AS Saldo, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo,MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta, Tasas.MontoCordobas FROM  Grupos INNER JOIN Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
                        "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
                        "GROUP BY Cuentas.TipoCuenta, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, Transacciones.CodCuentas, Tasas.MontoCordobas " & _
                        "HAVING  (Cuentas.TipoCuenta = 'Gastos') AND (SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 2)) - SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Debito, 2)) <> 0) ORDER BY Cuentas.KeyGrupo"
        ElseIf FrmReportes.OptPeriodo.Value = True Then
          SqlGastos = "SELECT Transacciones.CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta,AVG(Transacciones.TCambio) AS TCambio, SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Debito, 2)) AS Debito,SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 2)) AS Credito,SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Debito, 2))-SUM (Round(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 2)) AS Saldo, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo,MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta, Tasas.MontoCordobas FROM  Grupos INNER JOIN Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
                        "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
                        "GROUP BY Cuentas.TipoCuenta, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, Transacciones.CodCuentas, Tasas.MontoCordobas " & _
                        "HAVING (Cuentas.TipoCuenta = 'Gastos') AND (SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Credito, 2)) - SUM(ROUND(Transacciones.TCambio / Tasas.MontoCordobas * Transacciones.Debito, 2)) <> 0) ORDER BY Cuentas.KeyGrupo"
        End If
   End If
      
    Set Me.SrptDetalleGastos.object = New ArepResultadoDetalle
    Me.SrptDetalleGastos.object.LblTotalGrupo.Caption = "Total de Gastos"
    Me.SrptDetalleGastos.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptDetalleGastos.object.DataControl1.Source = SqlGastos
'    TotalGastos = Me.SrptDetalleIngresos.object.FldTotal.Text

 '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 '////////////////////////////////////TOTAL INGRESOS////////////////////////////////////////////////////////////////////////////////
 '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
     If FrmReportes.OptAcumulado.Value = True Then
        SqlIngresos = "SELECT MAX(Transacciones.CodCuentas) AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Credito) - SUM(Transacciones.TCambio * Transacciones.Debito) AS Saldo, MAX(Cuentas.KeyGrupo) AS KeyGrupo, MAX(Cuentas.DescripcionGrupo) AS DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta AS Expr1  " & _
                      "FROM  Grupos INNER JOIN Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo  " & _
                      "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102))  " & _
                      "GROUP BY Cuentas.TipoCuenta HAVING  (SUM(Transacciones.TCambio * Transacciones.Credito) - SUM(Transacciones.TCambio * Transacciones.Debito) <> 0) AND (Cuentas.TipoCuenta = N'Ingresos - Ventas') ORDER BY MAX(Cuentas.KeyGrupo)"
     ElseIf FrmReportes.OptPeriodo.Value = True Then
        SqlIngresos = "SELECT MAX(Transacciones.CodCuentas) AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Credito) - SUM(Transacciones.TCambio * Transacciones.Debito) AS Saldo, MAX(Cuentas.KeyGrupo) AS KeyGrupo, MAX(Cuentas.DescripcionGrupo) AS DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta AS Expr1  " & _
                      "FROM  Grupos INNER JOIN Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo  " & _
                      "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102))  " & _
                      "GROUP BY Cuentas.TipoCuenta HAVING  (SUM(Transacciones.TCambio * Transacciones.Credito) - SUM(Transacciones.TCambio * Transacciones.Debito) <> 0) AND (Cuentas.TipoCuenta = N'Ingresos - Ventas') ORDER BY MAX(Cuentas.KeyGrupo)"
     End If
        
        FrmReportes.DtaConsulta.RecordSource = SqlIngresos
        FrmReportes.DtaConsulta.Refresh
        If Not FrmReportes.DtaConsulta.Recordset.EOF Then
          Totalingresos = FrmReportes.DtaConsulta.Recordset("Saldo")
        End If
    
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////TOTAL COSTOS///////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
     If FrmReportes.OptAcumulado.Value = True Then
        SqlCostos = "SELECT MAX(Transacciones.CodCuentas) AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Debito)- SUM(Transacciones.TCambio * Transacciones.Credito) AS Saldo, MAX(Cuentas.KeyGrupo) AS KeyGrupo, MAX(Cuentas.DescripcionGrupo) AS DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta AS Expr1  " & _
                     "FROM  Grupos INNER JOIN Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo  " & _
                     "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102))  " & _
                     "GROUP BY Cuentas.TipoCuenta HAVING  (SUM(Transacciones.TCambio * Transacciones.Credito) - SUM(Transacciones.TCambio * Transacciones.Debito) <> 0) AND (Cuentas.TipoCuenta = N'Costos') ORDER BY MAX(Cuentas.KeyGrupo)"
     ElseIf FrmReportes.OptPeriodo.Value = True Then
         SqlCostos = "SELECT MAX(Transacciones.CodCuentas) AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Debito)- SUM(Transacciones.TCambio * Transacciones.Credito) AS Saldo, MAX(Cuentas.KeyGrupo) AS KeyGrupo, MAX(Cuentas.DescripcionGrupo) AS DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta AS Expr1  " & _
                     "FROM  Grupos INNER JOIN Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo  " & _
                     "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102))  " & _
                     "GROUP BY Cuentas.TipoCuenta HAVING  (SUM(Transacciones.TCambio * Transacciones.Credito) - SUM(Transacciones.TCambio * Transacciones.Debito) <> 0) AND (Cuentas.TipoCuenta = N'Costos') ORDER BY MAX(Cuentas.KeyGrupo)"
     End If
    
    FrmReportes.DtaConsulta.RecordSource = SqlCostos
    FrmReportes.DtaConsulta.Refresh
    If Not FrmReportes.DtaConsulta.Recordset.EOF Then
      TotalCostos = FrmReportes.DtaConsulta.Recordset("Saldo")
    End If
     
     '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
     '////////////////////////////////////TOTAL GASTOS//////////////////////////////////////////////////////////////////////////
     '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
     
     If FrmReportes.OptAcumulado.Value = True Then
            SqlGastos = "SELECT MAX(Transacciones.CodCuentas) AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Debito)- SUM(Transacciones.TCambio * Transacciones.Credito) AS Saldo, MAX(Cuentas.KeyGrupo) AS KeyGrupo, MAX(Cuentas.DescripcionGrupo) AS DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta AS Expr1  " & _
                        "FROM  Grupos INNER JOIN Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo  " & _
                        "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Fecha2 & "', 102))  " & _
                        "GROUP BY Cuentas.TipoCuenta HAVING  (SUM(Transacciones.TCambio * Transacciones.Credito) - SUM(Transacciones.TCambio * Transacciones.Debito) <> 0) AND (Cuentas.TipoCuenta = N'Gastos') ORDER BY MAX(Cuentas.KeyGrupo)"
     Else
            SqlGastos = "SELECT MAX(Transacciones.CodCuentas) AS CodCuentas, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.NombreCuenta) AS NombreCuenta, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.TCambio * Transacciones.Debito) AS Debito, SUM(Transacciones.TCambio * Transacciones.Credito) AS Credito, SUM(Transacciones.TCambio * Transacciones.Debito)- SUM(Transacciones.TCambio * Transacciones.Credito) AS Saldo, MAX(Cuentas.KeyGrupo) AS KeyGrupo, MAX(Cuentas.DescripcionGrupo) AS DescripcionGrupo, MAX(Cuentas.TipoMoneda) AS MonedaCuenta, Cuentas.TipoCuenta AS Expr1  " & _
                        "FROM  Grupos INNER JOIN Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo  " & _
                        "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102))  " & _
                        "GROUP BY Cuentas.TipoCuenta HAVING  (SUM(Transacciones.TCambio * Transacciones.Credito) - SUM(Transacciones.TCambio * Transacciones.Debito) <> 0) AND (Cuentas.TipoCuenta = N'Gastos') ORDER BY MAX(Cuentas.KeyGrupo)"
     End If
    
    FrmReportes.DtaConsulta.RecordSource = SqlGastos
    FrmReportes.DtaConsulta.Refresh
    If Not FrmReportes.DtaConsulta.Recordset.EOF Then
      TotalGastos = FrmReportes.DtaConsulta.Recordset("Saldo")
    End If
    
     '/////////////////////////////////////BUSCO LA UTILIDAD DEL PERIODO ////////////////////////
     MDIPrimero.AdoConsulta.RecordSource = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden, CodCuentas, Ubicacion, Orden AS Expr1, Haber3 - Debe3 AS Resultado From Reportes WHERE (KeyGrupo = 'RP') ORDER BY Expr1"
     MDIPrimero.AdoConsulta.Refresh
     If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
       UtilidadNeta = MDIPrimero.AdoConsulta.Recordset("Resultado")
  
     End If
    
'    UtilidadNeta = Totalingresos - TotalCostos - TotalGastos
    
    Me.LblUtilidadNeta.Caption = Format(UtilidadNeta, "##,##0.00")
    
End Sub

Private Sub ActiveReport_ReportEnd()
 On Error GoTo TipoErrs

If FrmReportes.ChkExportar.Value = 1 Then
Dim myExportObject As ActiveReportsExcelExport.ARExportExcel
Dim Nombre As String
    
Set myExportObject = CreateObject("ActiveReportsExcelExport.ARExportExcel")
'myExportObject.FileName = RutaArchivo
myExportObject.FileName = FrmReportes.CommonDialog1.FileName + ".xls"
myExportObject.Export Me.Pages
Set myExportObject = Nothing

MsgBox "Se ha Exportado con Exito!!!!"

End If

Exit Sub
TipoErrs:
    If err.Number <> 0 Then Exit Sub

End Sub

