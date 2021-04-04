VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepTotalAuxiliar 
   Caption         =   "Total de Reportes Auxiliares"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "ArepTotalAuxiliar.dsx":0000
End
Attribute VB_Name = "ArepTotalAuxiliar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClienteID As String
Private Sub ActiveReport_FetchData(EOF As Boolean)
If Not EOF Then
'Gets the current records SupplierID
ClienteID = DataControl1.Recordset.Fields("CodCuentas")
End If
End Sub
 Private Sub ActiveReport_hyperLink(ByVal Button As Integer, Link As String)
 Dim CodigoCuenta As String, FechaIni As Date, FechaFin As Date, SQL As String
 Dim rpt As Object, fPreview As New FrmPreview
'Check to see if an email link or web page has been selected
  If InStr(1, Link, "htm", vbTextCompare) = 0 And InStr(1, Link, "mailto", vbTextCompare) = 0 Then


    CodigoCuenta = Link
    
    FechaIni = FrmReportes.DTFecha1.Value
    FechaFin = FrmReportes.DTFecha2.Value
    
    
'    SQL = "SELECT Transacciones.CodCuentas, SUM(Transacciones.Debito * Transacciones.TCambio) AS Debito, SUM(Transacciones.Credito * Transacciones.TCambio) AS Credito, SUM((Transacciones.Debito - Transacciones.Credito) * Transacciones.TCambio) AS Saldo, SUM(Transacciones.TCambio) AS Expr5 " & _
'          "FROM  Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas " & _
'          "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) AND (Cuentas.TipoCuenta = 'Cuentas x Cobrar') GROUP BY Transacciones.CodCuentas HAVING (Transacciones.CodCuentas BETWEEN '" & CodigoCuenta & "' AND '" & CodigoCuenta & "')"
    
     SQL = "SELECT     Transacciones.CodCuentas AS CodCuentas, Transacciones.DescripcionMovimiento AS DescripcionMovimiento, " & _
         "Transacciones.FechaTransaccion AS FechaTransaccion, Transacciones.FechaVence AS FechaVence, " & _
         "Transacciones.NumeroMovimiento AS NumeroMovimiento, Cuentas.TipoCuenta AS TipoCuenta, Transacciones.FacturaNo AS FacturaNo, " & _
         "Transacciones.Debito * Transacciones.TCambio AS Debito, Transacciones.Credito * Transacciones.TCambio AS Credito, " & _
         "(Transacciones.Debito - Transacciones.Credito) * Transacciones.TCambio AS Saldo, Transacciones.ChequeNo, Cuentas.DescripcionCuentas, " & _
         "Transacciones.TCambio " & _
         "FROM         Transacciones INNER JOIN " & _
         "Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas " & _
         "WHERE     (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND ((Transacciones.FechaTransaccion) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "') " & _
         "ORDER BY Transacciones.FacturaNo,Transacciones.FechaTransaccion "
     
    
      ArepEstadoCuenta.DataControl1.ConnectionString = ConexionReporte
      ArepEstadoCuenta.DataControl1.Source = SQL
      ArepEstadoCuenta.LblFecha1.Caption = FrmReportes.DTFecha1.Value
      ArepEstadoCuenta.LblFecha.Caption = FrmReportes.DTFecha2.Value
    
       ArepEstadoCuenta.Logo.Picture = LoadPicture(RutaLogo)
    
      ArepEstadoCuenta.LblEmpresa.Caption = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
      ArepEstadoCuenta.LblEmpresa1.Caption = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
      ArepEstadoCuenta.LblEmpresa2.Caption = "RUC " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
'      ArepEstadoCuenta.Show 1
   
         Set rpt = New ArepEstadoCuenta
         rpt.DataControl1.ConnectionString = ConexionReporte
         rpt.DataControl1.Source = SQL
         fPreview.RunReport rpt
         fPreview.Show 1


  End If
End Sub

Private Sub ActiveReport_ReportStart()
 On Error GoTo err
 
 Select Case FrmReportes.CmbReportes.Text
      Case "LISTA CUENTAS X COBRAR": Me.Field22.Style = "font-size: 8.5pt; text-decoration: underline; color: rgb(0,0,255)"
      
 End Select
 
 
If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    End If
err:
    If err.Number <> 0 Then MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation
    
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

Private Sub Detail_Format()
Dim CodigoCuenta As String, TipoCuenta As String
Dim SaldoIni As Double, Debito As Double, Credito As Double
Dim KeyGrupo As String, SaldoFinal As Double

    CodigoCuenta = Me.Field22.Text
    FrmReportes.DtaConsulta.RecordSource = "SELECT  * From Cuentas WHERE (CodCuentas = '" & CodigoCuenta & "')"
    FrmReportes.DtaConsulta.Refresh
    If Not FrmReportes.DtaConsulta.Recordset.EOF Then
      TipoCuenta = FrmReportes.DtaConsulta.Recordset("TipoCuenta")
      If Me.Field21.Text <> "" Then
         SaldoIni = Me.Field21.Text
      End If
      If Me.Field18.Text <> "" Then
        Debito = Me.Field18.Text
      End If
      If Me.Field19.Text <> "" Then
        Credito = Me.Field19.Text
      End If
      
      If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
        Me.LblFinal.Caption = Format(SaldoIni + Debito - Credito, "##,##0.00")
      Else
        Me.LblFinal.Caption = Format(SaldoIni + Credito - Debito, "##,##0.00")
      End If
    
    Else
      If Me.Field25.Text <> "" Then
         KeyGrupo = Me.Field25.Text
         KeyGrupo = Mid(KeyGrupo, 1, 1)
         If Me.Field24.Text <> "" Then
          SaldoFinal = Me.Field24.Text
         Else
          SaldoFinal = 0
         End If

         If KeyGrupo = "A" Or KeyGrupo = "G" Or KeyGrupo = "O" Then
                Me.LblFinal.Caption = Format(SaldoFinal, "##,##0.00")
         Else
                Me.LblFinal.Caption = Format(SaldoFinal, "##,##0.00")
        End If
      End If
      
      
     
    End If
    
    Me.Field22.Hyperlink = ClienteID
End Sub

