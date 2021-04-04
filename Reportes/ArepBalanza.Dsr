VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepBalanza 
   Caption         =   "Balanza de Comprobacion"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19368
   SectionData     =   "ArepBalanza.dsx":0000
End
Attribute VB_Name = "ArepBalanza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalDebe3 As Double
Dim TotalHaber3 As Double
Dim TotalDebito As Double, TotalCredito As Double, TotalDebe1 As Double, TotalHaber1 As Double
Dim CodigoCuenta As String

Private Sub ActiveReport_Activate()
QuienReporte = Me.Name
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)
    If Not EOF Then
    'Gets the current records SupplierID
        If Not IsNull(Me.DataControl1.Recordset.Fields("CodCuentas")) Then
          CodigoCuenta = Me.DataControl1.Recordset.Fields("CodCuentas")
        Else
          CodigoCuenta = ""
        End If
    End If
End Sub



Private Sub ActiveReport_hyperLink(ByVal Button As Integer, Link As String)
'Dim CodigoCuentaDesde As String, CodigoCuentaHasta As String
'Dim rpt As Object
'Dim fPreview As New FrmPreview

'
'   If InStr(1, Link, "htm", vbTextCompare) = 0 And InStr(1, Link, "mailto", vbTextCompare) = 0 Then
'          ArepAuxiliar.LblRangoFecha = "Desde " & FrmReportes.DTFecha1.Value & " Hasta " & FrmReportes.DTFecha2.Value
'          ArepAuxiliar.LblFecha = Format(Now, "dd/mm/yyyy")
'          ArepAuxiliar.DataControl1.ConnectionString = ConexionReporte
'
'
'
'
'            CodigoCuentaDesde = LeeCadena(Link, 1)
'            CodigoCuentaHasta = CodigoCuentaDesde
'
'          SQL = "SELECT Transacciones.CodCuentas,  MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Transacciones.NPeriodo) AS NPeriodo,MAX(Transacciones.NTransaccion) AS NTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.VoucherNo) AS VoucherNo, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.Clave) AS Clave, SUM(Transacciones.Debito) AS Debito, SUM(Transacciones.Credito) AS Credito, MAX(Transacciones.FacturaNo) AS FacturaNo, MAX(Transacciones.ChequeNo) AS ChequeNo, MAX(Transacciones.Fuente) AS Fuente, MAX(Cuentas.TipoCuenta) AS TipoCuenta, SUM(Transacciones.Debito + Transacciones.Credito) As Saldo FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE  (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas " & _
'                                               "HAVING (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') AND (SUM(Transacciones.Debito + Transacciones.Credito) <> 0) ORDER BY Transacciones.CodCuentas"
'
''         SQL = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.TCambio, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTransaccion, Cuentas.TipoCuenta, Transacciones.TCambio AS Expr1, Transacciones.TCambio * Transacciones.Debito AS Debito, Transacciones.TCambio * Transacciones.Credito AS Credito FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
''                     "WHERE (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') AND (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME,'" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"
'               ArepAuxiliar.DataControl1.ConnectionString = ConexionReporte
'
'
'            Set rpt = New ArepAuxiliar
'            rpt.DataControl1.ConnectionString = ConexionReporte
'            rpt.DataControl1.Source = SQL
'            fPreview.RunReport rpt
'            fPreview.Show 1
'    End If


End Sub

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
On Error GoTo err

               QuienReporte = Me.Name

                Me.LblFechaFin = FrmReportes.DTFecha2.Value
                Me.LblFechaIni = FrmReportes.DTFecha1.Value
                Me.Field8.Visible = False
                Me.Field9.Visible = False
                Me.Field10.Visible = False
                Me.Field11.Visible = False
                Me.FldTDebe3.Visible = False
                Me.FldTHaber3.Visible = False
                Me.LblEmpresa = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
                Me.LblEmpresa1 = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
                Me.LblEmpresa2 = "RUC: " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
                Me.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
'                Me.LblRango.Caption = "Filtrado Desde: " & FrmReportes.CodDesde & " Hasta " & FrmReportes.CodHasta
    
                Me.FldCuentas.Hyperlink = ""

                 Me.LblMoneda.Caption = FrmReportes.CmbMoneda.Text

If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
End If
err:
    Exit Sub
End Sub

Private Sub Detail_Format()
Dim Nivel As Double, CodCuenta As String, TipoCuenta As String, Debito As Double, Credito As Double, Total1 As Double, Debito1 As Double, Credito1 As Double
Dim TipoGrupo As String

CodCuenta = Me.FldCuentas.Text
If FrmReportes.ChkExportar.Value = 0 Then
  Me.FldCuenta.Hyperlink = CodCuenta
  Me.FldCuenta.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 9pt"
End If


'/////////////////////////////////////////BUSCO EL TIPO DE LA CUENTA //////////////////////////////////////////////////
FrmReportes.DtaConsulta.RecordSource = "SELECT  * From Cuentas WHERE (CodCuentas = '" & CodCuenta & "')"
FrmReportes.DtaConsulta.Refresh
If Not FrmReportes.DtaConsulta.Recordset.EOF Then
 If Not IsNull(FrmReportes.DtaConsulta.Recordset("TipoCuenta")) Then
  TipoCuenta = FrmReportes.DtaConsulta.Recordset("TipoCuenta")
 Else
   MsgBox "Una cuenta no tiene definica el tipo de Cuenta", vbCritical, "Zeus Contable"
  End If
End If


If CodCuenta = "" Then
  '/////////////////////////////////////SI NO TIENE CODCUENTA SIGNIFICA QUE ES DE MAYOR ////////////////////////
  TipoGrupo = Mid(Me.FldKeyGrupo.Text, 1, 1)
  
  Select Case TipoGrupo
     Case "A": TipoCuenta = "Otros Activos"
     Case "B": TipoCuenta = "Pasivo"
     Case "C": TipoCuenta = "Capital"
     Case "D": TipoCuenta = "Ingresos - Ventas"
     Case "G": TipoCuenta = "Costos"
     Case "O": TipoCuenta = "Gastos"
  End Select
  

End If

     Debito1 = 0
     Credito1 = 0
     
     '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
     '/////////////////////////////////////BUSCO EL SALDO INCIAL DE LOS MOVIMIENTOS /////////////////////////////////////////////////
     '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
     If FrmReportes.CmbMoneda.Text = "Córdobas" Then
        FrmReportes.DtaConsulta.RecordSource = "SELECT  Cuentas.CodCuentas AS Cuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio - Transacciones.Credito * Transacciones.TCambio) AS Total, Cuentas.DescripcionCuentas,  Cuentas.TipoCuenta , Cuentas.TipoMoneda  FROM  Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
                                               "WHERE (Transacciones.FechaTransaccion < CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda HAVING  (Cuentas.CodCuentas = '" & CodCuenta & "') ORDER BY Cuentas.CodCuentas"
     Else
        FrmReportes.DtaConsulta.RecordSource = "SELECT  MAX(Cuentas.CodCuentas) AS CodCuentas, SUM(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas)) AS MDebito, SUM(Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas)) AS MCredito, SUM(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas) - Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NTransaccion) AS Transaccion FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas " & _
                                               "WHERE (Transacciones.FechaTransaccion < CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "', 102)) AND (Cuentas.CodCuentas = '" & CodCuenta & "')"
     End If
     
     FrmReportes.DtaConsulta.Refresh
     If Not FrmReportes.DtaConsulta.Recordset.EOF Then
        If Not IsNull(FrmReportes.DtaConsulta.Recordset("MDebito")) Then
           Debito1 = FrmReportes.DtaConsulta.Recordset("MDebito")
        End If
        If Not IsNull(FrmReportes.DtaConsulta.Recordset("MCredito")) Then
           Credito1 = FrmReportes.DtaConsulta.Recordset("MCredito")
        End If
     End If


     Debito = 0
     Credito = 0
     '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
     '/////////////////////////////////////BUSCO EL SALDO FINAL DE LOS MOVIMIENTOS /////////////////////////////////////////////////
     '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
     FrmReportes.DtaConsulta.RecordSource = "SELECT  Cuentas.CodCuentas AS Cuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio - Transacciones.Credito * Transacciones.TCambio) AS Total, Cuentas.DescripcionCuentas,  Cuentas.TipoCuenta , Cuentas.TipoMoneda  FROM  Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
                                            "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda HAVING  (Cuentas.CodCuentas = '" & CodCuenta & "') ORDER BY Cuentas.CodCuentas"
     FrmReportes.DtaConsulta.Refresh
     If Not FrmReportes.DtaConsulta.Recordset.EOF Then
        If Not IsNull(FrmReportes.DtaConsulta.Recordset("MDebito")) Then
           Debito = FrmReportes.DtaConsulta.Recordset("MDebito")
        End If
        If Not IsNull(FrmReportes.DtaConsulta.Recordset("MCredito")) Then
           Credito = FrmReportes.DtaConsulta.Recordset("MCredito")
        End If
     End If


'     If Me.fldDebe3.Text = "" Then Me.fldDebe3.Text = "0.00"
'     If Me.FldHaber3.Text = "" Then Me.FldHaber3.Text = "0.00"

     If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
        '//////////////////////////////////////////CALCULO EL SALDO FINAL ////////////////////////////////////////////////////////////
         If Debito1 > Credito1 Then
           Total1 = Debito1 - Credito1
           TotalDebe1 = Total1 + TotalDebe1
        Else
           Total1 = Credito1 - Debito1
           TotalHaber1 = Total1 + TotalHaber1
        End If


       '//////////////////////////////////////////CALCULO EL SALDO FINAL ////////////////////////////////////////////////////////////
         If Debito > Credito Then
           Total1 = CDbl(Debito) - CDbl(Credito)
           TotalDebito = Total1 + TotalDebito
            Me.FldDebe3.Text = CDbl(Me.FldDebeT3.Text) - CDbl(Me.FldHaberT3.Text)
            Me.FldHaber3.Text = ""
        Else
           Total1 = CDbl(Debito) - CDbl(Credito)
           TotalCredito = Total1 + TotalCredito
           Me.FldDebe3.Text = CDbl(Me.FldDebeT3.Text) - CDbl(Me.FldHaberT3.Text)
           Me.FldHaber3.Text = ""
        End If

     Else

         '//////////////////////////////////////////////CALCULO EL SALDO INICIAL /////////////////////////////////////////////////////////////////////////
           If Credito1 > Debito1 Then
              Total1 = Credito1 - Debito1
              TotalHaber1 = Total1 + TotalHaber1
    
           Else
              Total1 = Debito1 - Credito1
              TotalDebe1 = Total1 + TotalDebe1
           End If

            '/////////////////////////////////////////////////CALCULO EL SALDO FINAL /////////////////////////////////////////////////////
            If Credito > Debito Then
               Total1 = Credito - Debito
               TotalCredito = Total1 + TotalCredito
               Me.FldHaber3.Text = CDbl(Me.FldHaberT3.Text) - CDbl(Me.FldDebeT3.Text)
               Me.FldDebe3.Text = ""
            Else
               Total1 = Credito - Debito
               TotalDebito = Total1 + TotalDebito
               Me.FldHaber3.Text = CDbl(Me.FldHaberT3.Text) - CDbl(Me.FldDebeT3.Text)
               Me.FldDebe3.Text = ""
            End If
     End If
     


'sólo ver el debe o el haber resultado según sea la naturaleza de la cuenta
'MsgBox Mid(Me.FldCuenta.Text, 1, 1)
'If Me.fldDebe3.Text = "" Then Me.fldDebe3.Text = "0.00"
'If Me.FldHaber3.Text = "" Then Me.FldHaber3.Text = "0.00"
'
'If Mid(Me.FldCuenta.Text, 1, 1) = "1" Or Mid(Me.FldCuenta.Text, 1, 1) = "5" Then
'   If CDbl(Me.fldDebe3.Text) > CDbl(Me.FldHaber3.Text) Then
'    Me.fldDebe3.Text = CDbl(Me.fldDebe3.Text) - CDbl(Me.FldHaber3.Text)
'    Me.FldHaber3.Text = "0.00"
'   Else
'    Me.FldHaber3.Text = CDbl(Me.FldHaber3.Text) - CDbl(Me.fldDebe3.Text)
'    Me.fldDebe3.Text = "0.00"
'   End If
'ElseIf Mid(Me.FldCuenta.Text, 1, 1) = "2" Or Mid(Me.FldCuenta.Text, 1, 1) = "3" Or Mid(Me.FldCuenta.Text, 1, 1) = "4" Then
'  If CDbl(Me.FldHaber3.Text) > CDbl(Me.fldDebe3.Text) Then
'    Me.FldHaber3.Text = CDbl(Me.FldHaber3.Text) - CDbl(Me.fldDebe3.Text)
'    Me.fldDebe3.Text = "0.00"
'  Else
'    Me.fldDebe3.Text = CDbl(Me.fldDebe3.Text) - CDbl(Me.FldHaber3.Text)
'    Me.FldHaber3.Text = "0.00"
'  End If
'End If

'TotalDebe3 = TotalDebe3 + CDbl(Me.fldDebe3.Text)
'TotalHaber3 = TotalHaber3 + CDbl(Me.FldHaber3.Text)

TotalDebe3 = TotalDebito
TotalHaber3 = TotalCredito

Me.FldDebe3.Text = Format(Me.FldDebe3, "###,##0.00")
Me.FldHaber3.Text = Format(Me.FldHaber3, "###,##0.00")

     If Me.FldDebe1.Text = "0.00" Then
       Me.FldDebe1.Text = ""
     End If
     
     If Me.fldDebe2.Text = "0.00" Then
       Me.fldDebe2.Text = ""
     End If
     
     If Me.FldDebe3.Text = "0.00" Then
       Me.FldDebe3.Text = ""
     End If
     
     If Me.FldHaber1.Text = "0.00" Then
       Me.FldHaber1.Text = ""
     End If

    If Me.FldHaber2.Text = "0.00" Then
       Me.FldHaber2.Text = ""
    End If

     If Me.FldHaber3.Text = "0.00" Then
       Me.FldHaber3.Text = ""
     End If

 If Me.FldNivel.Text <> "" Then
  Nivel = Me.FldNivel.Text
 Else
  Nivel = 0
 End If
 
 
      If CodCuenta = "" Then
         Select Case Nivel
                 Case 1
                        Me.FldCuenta.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 9pt; "
                        Me.FldDebe1.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                        Me.FldHaber1.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                        Me.fldDebe2.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                        Me.FldHaber2.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                        Me.FldDebe3.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                        Me.FldHaber3.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                        Me.FldDebe1.Alignment = ddTXRight
                        Me.FldHaber1.Alignment = ddTXRight
                        Me.fldDebe2.Alignment = ddTXRight
                        Me.FldHaber2.Alignment = ddTXRight
                        Me.FldDebe3.Alignment = ddTXRight
                        Me.FldHaber3.Alignment = ddTXRight
        
                 Case 2
                     Me.FldCuenta.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 9pt; "
                     Me.FldDebe1.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                     Me.FldHaber1.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                     Me.fldDebe2.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                     Me.FldHaber2.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                     Me.FldDebe3.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                     Me.FldHaber3.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                        Me.FldDebe1.Alignment = ddTXRight
                        Me.FldHaber1.Alignment = ddTXRight
                        Me.fldDebe2.Alignment = ddTXRight
                        Me.FldHaber2.Alignment = ddTXRight
                        Me.FldDebe3.Alignment = ddTXRight
                        Me.FldHaber3.Alignment = ddTXRight
                 Case 3
                     Me.FldCuenta.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 9pt; "
                     Me.FldDebe1.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                     Me.FldHaber1.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                     Me.fldDebe2.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                     Me.FldHaber2.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                     Me.FldDebe3.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                     Me.FldHaber3.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                        Me.FldDebe1.Alignment = ddTXRight
                        Me.FldHaber1.Alignment = ddTXRight
                        Me.fldDebe2.Alignment = ddTXRight
                        Me.FldHaber2.Alignment = ddTXRight
                        Me.FldDebe3.Alignment = ddTXRight
                        Me.FldHaber3.Alignment = ddTXRight
                 Case 4
'                     Me.FldCuenta.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 9pt; "
'                     Me.fldDebe1.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
'                     Me.FldHaber1.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
'                     Me.fldDebe2.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
'                     Me.FldHaber2.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
'                     Me.fldDebe3.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
'                     Me.FldHaber3.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
         Case Else
        '        Me.FldCuenta.Style = "color: rgb(0,0,0); font-weight: (null); font-size: 9pt"
                 Me.FldCuenta.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 9pt"
                     Me.FldDebe1.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 8pt; "
                     Me.FldHaber1.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 8pt; "
                     Me.fldDebe2.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 8pt; "
                     Me.FldHaber2.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 8pt; "
                     Me.FldDebe3.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 8pt; "
                     Me.FldHaber3.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 8pt; "
                        Me.FldDebe1.Alignment = ddTXRight
                        Me.FldHaber1.Alignment = ddTXRight
                        Me.fldDebe2.Alignment = ddTXRight
                        Me.FldHaber2.Alignment = ddTXRight
                        Me.FldDebe3.Alignment = ddTXRight
                        Me.FldHaber3.Alignment = ddTXRight
           
         End Select
      Else
                 Me.FldCuenta.Style = "color: rgb(0,0,128); font-weight: Arial Narrow; font-size: 8pt"
                     Me.FldDebe1.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 8pt; "
                     Me.FldHaber1.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 8pt; "
                     Me.fldDebe2.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 8pt; "
                     Me.FldHaber2.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 8pt; "
                     Me.FldDebe3.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 8pt; "
                     Me.FldHaber3.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 8pt; "
                        Me.FldDebe1.Alignment = ddTXRight
                        Me.FldHaber1.Alignment = ddTXRight
                        Me.fldDebe2.Alignment = ddTXRight
                        Me.FldHaber2.Alignment = ddTXRight
                        Me.FldDebe3.Alignment = ddTXRight
                        Me.FldHaber3.Alignment = ddTXRight
      End If

End Sub

Private Sub GroupFooter1_Format()
    Dim Debe1 As Double, Haber1 As Double, Debe2 As Double, Haber2 As Double, Debe3 As Double, Haber3 As Double
    Dim CodigoCuentaDesde As String, CodigoCuentaHasta As String
    
    Me.Field8.Visible = False
    Me.Field9.Visible = False
    Me.Field10.Visible = False
    Me.Field11.Visible = False
   
    '////////////////////////////TOTAL SALDO INICIAL //////////////////////////////////////////
'    FrmReportes.DtaConsulta.RecordSource = "SELECT MAX(Cuentas.CodCuentas) AS Cuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio - Transacciones.Credito * Transacciones.TCambio) AS Total, MAX(Cuentas.DescripcionCuentas)  AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda  FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
'                                           "WHERE (Transacciones.FechaTransaccion < '" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "') ORDER BY MAX(Cuentas.CodCuentas)"
'    FrmReportes.DtaConsulta.Refresh
'    If Not FrmReportes.DtaConsulta.Recordset.EOF Then
'       If Not IsNull(FrmReportes.DtaConsulta.Recordset("MDebito")) Then
'         Debe1 = FrmReportes.DtaConsulta.Recordset("MDebito")
'       End If
'       If Not IsNull(FrmReportes.DtaConsulta.Recordset("MCredito")) Then
'         Haber1 = FrmReportes.DtaConsulta.Recordset("MCredito")
'       End If
'    End If
                                          
     '///////////////////////////TOTAL MOVIMIENTOS ////////////////////////////////////////////
     If QUIEN = "Balanza" Then
     
         If FrmReportes.CmbMoneda.Text = "Córdobas" Then
            Ajuste = "Dólares"
         ElseIf FrmReportes.CmbMoneda.Text = "Dólares" Then
            Ajuste = "Córdobas"
         
         End If
     
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
     
     
         If FrmReportes.CmbMoneda.Text = "Córdobas" Then
           FrmReportes.DtaConsulta.RecordSource = "SELECT  MAX(Cuentas.CodCuentas) AS Expr1, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio - Transacciones.TCambio * Transacciones.Credito) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NTransaccion) AS Transaccion FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
                                                  "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) AND (Cuentas.KeyGrupo BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') AND (IndiceTransaccion.Ajuste <> 'Dólares')"
         Else
           FrmReportes.DtaConsulta.RecordSource = "SELECT  MAX(Cuentas.CodCuentas) AS Expr1, SUM(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas)) AS MDebito, SUM(Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas)) AS MCredito, SUM(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas) - Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NTransaccion) As Transaccion FROM   Cuentas INNER JOIN  Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN  Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas INNER JOIN  IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion " & _
                                                  "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) AND (Cuentas.KeyGrupo BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') AND (IndiceTransaccion.Ajuste <> 'Córdobas')"
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
     
         If FrmReportes.CmbMoneda.Text = "Córdobas" Then
           FrmReportes.DtaConsulta.RecordSource = "SELECT MAX(Descripcion) AS Descripcion, SUM(Debe1) AS Debe1, SUM(Haber1) AS Haber1, SUM(Debe2) AS MDebito, SUM(Haber2) AS MCredito, SUM(Debe3) AS Debe3, SUM(Haber3) As Haber3 From Reportes"
         Else
'           FrmReportes.DtaConsulta.RecordSource = "SELECT  MAX(Cuentas.CodCuentas) AS Expr1, SUM(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas)) AS MDebito, SUM(Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas)) AS MCredito, SUM(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas) - Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NTransaccion) As Transaccion FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
'                                                  "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) HAVING (MAX(Cuentas.CodCuentas) BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "')"
           FrmReportes.DtaConsulta.RecordSource = "SELECT MAX(Descripcion) AS Descripcion, SUM(Debe1) AS Debe1, SUM(Haber1) AS Haber1, SUM(Debe2) AS MDebito, SUM(Haber2) AS MCredito, SUM(Debe3) AS Debe3, SUM(Haber3) As Haber3 From Reportes"
         End If
    Else
    
         If FrmReportes.CmbMoneda.Text = "Córdobas" Then
           FrmReportes.DtaConsulta.RecordSource = "SELECT  MAX(Cuentas.CodCuentas) AS Expr1, SUM(Transacciones.Debito * Transacciones.TCambio) AS MDebito, SUM(Transacciones.TCambio * Transacciones.Credito) AS MCredito, SUM(Transacciones.Debito * Transacciones.TCambio - Transacciones.TCambio * Transacciones.Credito) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NTransaccion) AS Transaccion FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
                                                  "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) HAVING (MAX(Cuentas.CodCuentas) BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "')"
         Else
           FrmReportes.DtaConsulta.RecordSource = "SELECT  MAX(Cuentas.CodCuentas) AS Expr1, SUM(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas)) AS MDebito, SUM(Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas)) AS MCredito, SUM(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas) - Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas)) AS Total, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NTransaccion) As Transaccion FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
                                                  "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) HAVING (MAX(Cuentas.CodCuentas) BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "')"
         End If
    
    End If
    
    
    
      FrmReportes.DtaConsulta.Refresh
        If Not FrmReportes.DtaConsulta.Recordset.EOF Then
           If Not IsNull(FrmReportes.DtaConsulta.Recordset("MDebito")) Then
             Debe2 = FrmReportes.DtaConsulta.Recordset("MDebito")
           End If
           If Not IsNull(FrmReportes.DtaConsulta.Recordset("MCredito")) Then
             Haber2 = FrmReportes.DtaConsulta.Recordset("MCredito")
           End If
        End If
    
    Me.LblTotalDebe1.Caption = Format(TotalDebe1, "###,##0.00")
    Me.LblTotalHaber1.Caption = Format(TotalHaber1, "###,##0.00")
    Me.LblTotalDebe2.Caption = Format(Debe2, "###,##0.00")
    Me.LblTotalHaber2.Caption = Format(Haber2, "###,##0.00")
    Me.LblDebe3.Caption = Format(TotalDebe1 + Debe2, "###,##0.00")
    Me.LblHaber3.Caption = Format(TotalHaber1 + Haber2, "###,##0.00")
    
    
    '////////////////////////////////AHORA BUSCO LOS SALDO FINALES ////////////////////////////////////////////////////////
    
'    FrmReportes.DtaConsulta.RecordSource = "SELECT SUM(Debe3) AS Debe3, SUM(Haber3) AS Haber3 From Reportes Where (Not (CodCuentas Is Null))"
'    FrmReportes.DtaConsulta.Refresh
'    If FrmReportes.DtaConsulta.Recordset.EOF Then
'      Me.LblDebe3.Caption = Format(0, "###,##0.00")
'      Me.LblHaber3.Caption = Format(0, "###,##0.00")
'    Else
'       Me.LblDebe3.Caption = Format(FrmReportes.DtaConsulta.Recordset("Debe3"), "###,##0.00")
'      Me.LblHaber3.Caption = Format(FrmReportes.DtaConsulta.Recordset("Haber3"), "###,##0.00")
'
'    End If
    
    
'   If FrmReportes.TxtDesde.Text <> "" And FrmReportes.TxtHasta.Text <> "" Then
'    Me.LblDebe3.Caption = Format(TotalDebe3, "###,##0.00")
'    Me.LblHaber3.Caption = Format(TotalHaber3, "###,##0.00")
'   Else
'    Me.LblDebe3.Caption = Format(TotalDebe1 + Debe2, "###,##0.00")
'    Me.LblHaber3.Caption = Format(TotalHaber1 + Haber2, "###,##0.00")
'   End If

End Sub

