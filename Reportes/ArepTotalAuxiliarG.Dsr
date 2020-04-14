VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepTotalAuxiliarG 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   33655
   _ExtentY        =   20267
   SectionData     =   "ArepTotalAuxiliarG.dsx":0000
End
Attribute VB_Name = "ArepTotalAuxiliarG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Detail_Format()
  Dim Movimiento As Double
        CodigoCuenta = Me.Field16.Text
        FrmReportes.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
       'Me.DtaConsulta.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.VoucherNo, Transacciones.ChequeNo, Transacciones.DescripcionMovimiento, Transacciones.CodCuentas, Transacciones.Debito, Transacciones.Credito, 5000+[Transacciones]![Debito]-[Transacciones]![Credito] AS Balance, Transacciones.NPeriodo From Transacciones WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Transacciones.CodCuentas)='" & CodigoCuenta & "'))"
       FrmReportes.DtaConsulta.Refresh
       If Not FrmReportes.DtaConsulta.Recordset.EOF Then
         Periodo1 = FrmReportes.DtaConsulta.Recordset("NPeriodo")
         Periodo1 = Periodo1 - 1
        NumFecha1 = FrmReportes.DTFecha1.Value
        NumFecha2 = FrmReportes.DTFecha2.Value
'///////////////Busco el Acumulado de la cuenta hasta la ultima fecha Seleccionada////////////
'///////////////////////////////////////////////////////////////////////////////////////////////
         FrmReportes.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito From Transacciones Where (((Transacciones.FechaTransaccion) < " & NumFecha1 & ")) GROUP BY Transacciones.CodCuentas Having (((Transacciones.CodCuentas) = '" & CodigoCuenta & "'))"
         FrmReportes.DtaHistorial.Refresh
         
         If Not FrmReportes.DtaHistorial.Recordset.EOF Then
          If Not IsNull(FrmReportes.DtaHistorial.Recordset("MDebito")) Then
           Debito = FrmReportes.DtaHistorial.Recordset("MDebito")
           
          End If
           If Not IsNull(FrmReportes.DtaHistorial.Recordset("MCredito")) Then
             Credito = FrmReportes.DtaHistorial.Recordset("MCredito")
          
          End If
          Total = Debito - Credito
          SaldoIni = Total
         Else
          SaldoIni = 0
         End If
          
           
         Else
'///////////////Busco el Acumulado de la cuenta hasta la ultima fecha Seleccionada////////////
'///////////////////////////////////////////////////////////////////////////////////////////////
         FrmReportes.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito From Transacciones Where (((Transacciones.FechaTransaccion) < " & NumFecha1 & ")) GROUP BY Transacciones.CodCuentas Having (((Transacciones.CodCuentas) = '" & CodigoCuenta & "'))"
         FrmReportes.DtaHistorial.Refresh
         
         If Not FrmReportes.DtaHistorial.Recordset.EOF Then
          If Not IsNull(FrmReportes.DtaHistorial.Recordset("MDebito")) Then
           Debito = FrmReportes.DtaHistorial.Recordset("MDebito")
         
          End If
           If Not IsNull(FrmReportes.DtaHistorial.Recordset("MCredito")) Then
             Credito = FrmReportes.DtaHistorial.Recordset("MCredito")
        
          End If
          Total = Debito - Credito
          SaldoIni = Total
         Else
          SaldoIni = 0
         End If
         End If
         
 '//////////////////////////Busco el total de los movimientos///////////////////////////////
 
 TotalIni = SaldoIni + TotalIni
        
  Me.LblIni.Caption = Format(SaldoIni, "##,##0.00")
  If Me.Field20.Text = "" Then
   Movimiento = 0
  Else
   Movimiento = Me.Field20.Text
  End If
  FrmReportes.DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito, Sum(Debito*TCambio)-Sum(Credito*TCambio) AS Total, Cuentas.DescripcionCuentas FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE (((Transacciones.FechaTransaccion) Between " & NumFecha1 & " And " & NumFecha2 & ")) GROUP BY Cuentas.CodCuentas, Cuentas.DescripcionCuentas"
  FrmReportes.DtaConsulta.Refresh
   If FrmReportes.DtaConsulta.Recordset.EOF Then
         TotalFin = SaldoIni + TotalFin
     Me.LblFinal.Caption = Format(SaldoIni, "##,##0.00")
   Else
    TotalFin = SaldoIni + Movimiento + TotalFin
    Me.LblFinal.Caption = Format(Movimiento + SaldoIni, "##,##0.00")
   End If

End Sub

Private Sub GroupFooter1_Format()

'///////////////Busco el Acumulado de la cuenta hasta la ultima fecha Seleccionada////////////
'///////////////////////////////////////////////////////////////////////////////////////////////
    NumFecha1 = FrmReportes.DTFecha1.Value
    NumFecha2 = FrmReportes.DTFecha2.Value
    If FrmReportes.DBCodigo.Text = "" Then
     CodigoCuenta = Me.Field16.Text
    Else
     CodigoCuenta = FrmReportes.DBCodigo.Text
    End If
        FrmReportes.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito From Transacciones Where (((Transacciones.FechaTransaccion) <= " & NumFecha2 & ")) GROUP BY Transacciones.CodCuentas Having (((Transacciones.CodCuentas) = '" & CodigoCuenta & "'))"
        FrmReportes.DtaHistorial.Refresh
         
         If Not FrmReportes.DtaHistorial.Recordset.EOF Then
          If Not IsNull(FrmReportes.DtaHistorial.Recordset("MDebito")) Then
           Debito = FrmReportes.DtaHistorial.Recordset("MDebito")
          End If
           If Not IsNull(FrmReportes.DtaHistorial.Recordset("MCredito")) Then
             Credito = FrmReportes.DtaHistorial.Recordset("MCredito")
          End If
          Total = Debito - Credito
          SaldoFin = Total
                
           
         Else
           SaldoFin = 0
         End If



 'SaldoFin = SaldoIni + SaldoFin
 Me.LblFinal = Format(SaldoFin, "##,##0.00")
 Me.LblTotalFin.Caption = Format(TotalFin, "##,##0.00")
 Me.LblTotalIni.Caption = Format(TotalIni, "##,##0.00")
End Sub

Private Sub GroupHeader1_Format()
TotalIni = 0
TotalFin = 0
'///////////////Busco el Acumulado de la cuenta hasta la ultima fecha Seleccionada////////////
'///////////////////////////////////////////////////////////////////////////////////////////////
    NumFecha1 = FrmReportes.DTFecha1.Value
    NumFecha2 = FrmReportes.DTFecha2.Value
    If FrmReportes.DBCodigo.Text = "" Then
     CodigoCuenta = Me.Field16.Text
    Else
     CodigoCuenta = FrmReportes.DBCodigo.Text
    End If
        FrmReportes.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito From Transacciones Where (((Transacciones.FechaTransaccion) < " & NumFecha1 & ")) GROUP BY Transacciones.CodCuentas Having (((Transacciones.CodCuentas) = '" & CodigoCuenta & "'))"
        FrmReportes.DtaHistorial.Refresh
         
         If Not FrmReportes.DtaHistorial.Recordset.EOF Then
          If Not IsNull(FrmReportes.DtaHistorial.Recordset("MDebito")) Then
           Debito = FrmReportes.DtaHistorial.Recordset("MDebito")
          End If
           If Not IsNull(FrmReportes.DtaHistorial.Recordset("MCredito")) Then
             Credito = FrmReportes.DtaHistorial.Recordset("MCredito")
          End If
          Total = Debito - Credito
          SaldoIni = Total
                
           
         Else
           SaldoIni = 0
         End If
         

     Me.LblIni.Caption = Format(SaldoIni, "##,##0.00")
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
    FrmReportes.CDRuta.flags = cdlOFNHideReadOnly
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


