VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepBank 
   Caption         =   "REPORTES DFID"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ActiveReport1.dsx":0000
End
Attribute VB_Name = "ArepBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalDebito As Double, TotalCredito As Double
 
Private Sub ActiveReport_ReportEnd()
 On Error GoTo err:
   Unload SubDetalle.object
   Set SubDetalle.object = Nothing
   
   Unload Me.SubFlotantes.object
   Set Me.SubFlotantes.object = Nothing
   
   Unload Me.SubCheques.object
   Set Me.SubCheques.object = Nothing
   

   
  
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

Set SubDetalle.object = New ArepBakControlSub
Set Me.SubFlotantes.object = New ArepSptDepositos
Set Me.SubCheques.object = New ArepSptDepositos

TotalConDebito = 0
TotalConCredito = 0



       
      
       
 On Error GoTo err
If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    End If
err:
    If err.Number <> 0 Then MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation
    
End Sub

Private Sub Detail_Format()
Dim Mov1 As Double, Mov2 As Double
Dim NumeroMovimiento As Integer

 If Me.Field5.Text = "0.00" Or Me.Field5.Text = "" Then
   Mov1 = 0
 Else
   Mov1 = Me.Field5.Text
 End If
 
  If Me.Field6.Text = "0.00" Or Me.Field6.Text = "" Then
   Mov2 = 0
 Else
   Mov2 = Me.Field6.Text
 End If
If Not Mov1 = 0 Then
 SaldoFin = Mov1 - Mov2 + SaldoFin
End If
If Not Mov2 = 0 Then
 SaldoFin = Mov1 - Mov2 + SaldoFin
End If
Me.LblBalances = Format(SaldoFin, "##,##0.00")
If Me.Field10.Text = "" Then
 NumeroMovimiento = -1
Else
 NumeroMovimiento = Int(Me.Field10.Text)
End If
If Me.Field1.Text = "" Then
  Fecha = Format(Now, "dd/mm/yyyy")
Else
Fecha = Me.Field1.Text
End If

TotalDebito = Mov1 + TotalDebito
TotalCredito = Mov2 + TotalCredito

NumFecha1 = Fecha
NumFecha2 = Fecha

SubDetalle.object.DataControl1.ConnectionString = ConexionReporte
'SubDetalle.object.DataControl1.Source = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.Debito, Transacciones.Credito, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTransaccion FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.NumeroMovimiento)=" & NumeroMovimiento & ") AND ((Transacciones.FechaTransaccion) Between " & NumFecha1 & "  And " & NumFecha2 & ")) ORDER BY Transacciones.NTransaccion "
SubDetalle.object.DataControl1.Source = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.Debito, Transacciones.Credito, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTransaccion, Transacciones.Beneficiario FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.CodCuentas)<>'" & CodigoBanco & "') AND ((Transacciones.NumeroMovimiento)=" & NumeroMovimiento & ") AND ((Transacciones.FechaTransaccion) Between " & NumFecha1 & "  And " & NumFecha2 & ")) ORDER BY Transacciones.NTransaccion"
'SubDetalle.object.DataControl1.Source = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.NumeroMovimiento,Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.Debito, Transacciones.Credito,Transacciones.FacturaNo , Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTransaccion, Transacciones.Beneficiario FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE     (Transacciones.CodCuentas <> '" & CodigoBanco & "') AND ((Transacciones.NumeroMovimiento)=" & NumeroMovimiento & ") AND (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(Fecha, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Fecha, "yyyy-mm-dd") & "', 102)) ORDER BY Transacciones.NTransaccion"


End Sub

Private Sub GroupFooter1_BeforePrint()
   Me.LblTotalDepositos.Caption = Format(TotalConDebito, "##,##0.00")
   Me.LblTotalCheques.Caption = Format(TotalConCredito, "##,##0.00")
End Sub

Private Sub GroupFooter1_Format()
Dim Total1 As Double, Total2 As Double
Dim CodigoCuenta As String, Fecha As String

Fecha = Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd")
CodigoCuenta = FrmReportes.DBCodigo.Text
Me.LblCuentaBanco.Caption = CodigoCuenta




Me.LblCloseBalance.Caption = Format(SaldoFin, "##,##0.00")
Me.LblSaldoFinal.Caption = Format(SaldoFin, "##,##0.00")
Me.LblAnterior.Caption = Format(SaldoIni, "##,##0.00")
Me.LblIngresos.Caption = Format(TotalDebito, "##,##0.00")
Me.LblEgresos.Caption = Format(TotalCredito, "##,##0.00")

Total1 = Me.Field12.Text
If Me.LblPaidIn.Caption <> "" Then
 Total2 = Me.LblPaidIn.Caption
Else
 Total2 = 0
End If
Me.LblTDebito.Caption = Format(Total1 + Total2, "##,##0.00")
If CodigoCuenta <> "" Then
    Me.SubFlotantes.object.DataControl1.ConnectionString = ConexionReporte
    Me.SubFlotantes.object.DataControl1.Source = "SELECT FechaTransaccion, DescripcionMovimiento, ChequeNo, VoucherNo, TCambio, Debito * TCambio AS Debito, Conciliada, TCambio * Credito AS Credito, ConciliacionProcesada From Transacciones  " & _
                                            "WHERE (CodCuentas = '" & CodigoCuenta & "') AND (NombreCuenta <> '**********CANCELADO*************') AND (FechaTransaccion <= CONVERT(DATETIME, '" & Fecha & "', 102)) AND (Conciliada <> 1) AND (TCambio * Debito <> 0) ORDER BY FechaTransaccion, NumeroMovimiento"
    Me.SubFlotantes.object.FldDebito.Visible = True
    Me.SubFlotantes.object.FldCredito.Visible = False
    

    Me.SubCheques.object.DataControl1.ConnectionString = ConexionReporte
    Me.SubCheques.object.DataControl1.Source = "SELECT FechaTransaccion, DescripcionMovimiento, ChequeNo, VoucherNo, TCambio, Debito * TCambio AS Debito, Conciliada, TCambio * Credito AS Credito, ConciliacionProcesada From Transacciones  " & _
                                            "WHERE (CodCuentas = '" & CodigoCuenta & "') AND (NombreCuenta <> '**********CANCELADO*************') AND (FechaTransaccion <= CONVERT(DATETIME, '" & Fecha & "', 102)) AND (Conciliada <> 1) AND (TCambio * Credito <> 0) ORDER BY FechaTransaccion, NumeroMovimiento"
    Me.SubCheques.object.FldDebito.Visible = False
    Me.SubCheques.object.FldCredito.Visible = True


End If


End Sub

Private Sub GroupHeader1_Format()
CodigoBanco = Me.Field11
SaldoFin = SaldoIni

End Sub

Private Sub PageHeader_Format()

    If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    Else
'        MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation
    End If
'       Me.Logo.Picture = LoadPicture(RutaLogo)
       Me.LblEmpresa = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
       Me.LblEmpresa1 = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
       Me.LblEmpresa2 = "RUC: " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
       Me.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
       Me.LblMoneda.Caption = FrmReportes.CmbMoneda.Text

       Me.LblBalance = Format(SaldoIniBank, "##,##0.00")
       Me.LblPaidIn.Caption = Format(SaldoIniBank, "##,##0.00")
       Me.LblTipo.Caption = "CUENTA DE BANCO: " & CodigoBancoBank
       Me.LblFecha1 = FechaIniBank
       Me.LblFecha2 = FechaFinBank
End Sub
