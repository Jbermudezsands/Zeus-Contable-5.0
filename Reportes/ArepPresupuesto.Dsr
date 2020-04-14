VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepPresupuesto 
   Caption         =   "Reporte de Presupuesto"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "ArepPresupuesto.dsx":0000
End
Attribute VB_Name = "ArepPresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
 On Error GoTo err
 
     Me.Label13.Caption = "PRESUPUESTO PARA EL AÑO " & Year(Fecha1)
     Me.Logo.Picture = LoadPicture(RutaLogo)
     Me.LblEmpresa = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
     Me.LblEmpresa1 = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
     Me.LblEmpresa2 = "RUC: " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
     Me.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
 
If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    End If
err:
    If err.Number <> 0 Then MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation
    End Sub

Private Sub Detail_Format()

 
 Criterio = "CodCuentas='" & Me.Field2.Text & "'"
FrmReportes.DtaCuentas.Recordset.Find (Criterio)
If Not FrmReportes.DtaCuentas.Recordset.EOF Then

    TipoMoneda = FrmReportes.DtaCuentas.Recordset("TipoMoneda")
   Select Case TipoMoneda
      Case "Dólares"
         Fecha = Format(Now, "DD/MM/YYYY")
         NumFecha1 = Fecha
         FrmReportes.DtaConsulta.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) = " & NumFecha1 & "))"
         FrmReportes.DtaConsulta.Refresh
         If Not FrmReportes.DtaConsulta.Recordset.EOF Then
           MontoTasa = FrmReportes.DtaConsulta.Recordset("MontoLibras")
         End If
      Case "Libras"
         MontoTasa = 1
   End Select

  CodigoCuenta = Me.Field2.Text
  FrmReportes.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito From Transacciones GROUP BY Transacciones.CodCuentas Having (((Transacciones.CodCuentas) = '" & CodigoCuenta & "'))"
  FrmReportes.DtaConsulta.Refresh
  If Not FrmReportes.DtaConsulta.Recordset.EOF Then
      If Not IsNull(FrmReportes.DtaConsulta.Recordset("MDebito")) Then
       Debito = FrmReportes.DtaConsulta.Recordset("MDebito")
      End If
      If Not IsNull(FrmReportes.DtaConsulta.Recordset("MCredito")) Then
       Credito = FrmReportes.DtaConsulta.Recordset("MCredito")
      End If
   Select Case TipoMoneda
     Case "Dólares"
        If Not MontoTasa = 0 Then
         Saldo = (Debito - Credito) / MontoTasa
        Else
         Saldo = 0
        End If
     Case "Libras"
         Saldo = (Debito - Credito)
     Case "Córdobas"
         Saldo = (Debito - Credito)
   End Select
   Me.LblReal.Caption = Format(Saldo, "##,##0.00")
   Total = Total + Saldo
   Presupuesto = Me.Field4
   
   Me.LblDiferencia.Caption = Format(Presupuesto - Saldo, "##,##0.00")
  If Not Saldo = 0 And Not Presupuesto = 0 Then
   Me.LblPorciento.Caption = Format((Saldo) / (Presupuesto), "0.00%")
  End If
   End If
  
 End If
 
End Sub

Private Sub GroupFooter1_Format()
 Presupuesto = Me.Field6
 Me.Label11.Caption = Format(Total, "##,##0.00")
 Me.LblDiferencia2.Caption = Format(Presupuesto - Total, "##,##0.00")
 If Not Total = 0 And Not Presupuesto = 0 Then
  Me.LblPorciento2.Caption = Format((Total) / (Presupuesto), "0.00%")
 End If
 SaldoFin = SaldoFin + Total
 
 Presupuesto = Me.Field6
 Me.LblSaldoFin.Caption = Format(SaldoFin, "##,##0.00")
 Me.LblDiferencia3.Caption = Format(Presupuesto - SaldoFin, "##,##0.00")
If Not SaldoFin = 0 And Not Presupuesto = 0 Then
 Me.LblPorciento3.Caption = Format((SaldoFin) / (Presupuesto), "0.00%")
End If

End Sub

Private Sub GroupHeader1_Format()
Total = 0
Saldo = 0
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

