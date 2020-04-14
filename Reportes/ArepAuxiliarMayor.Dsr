VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepAuxiliarMayor 
   Caption         =   "Reporte de Analitico"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "ArepAuxiliarMayor.dsx":0000
End
Attribute VB_Name = "ArepAuxiliarMayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub ActiveReport_ReportStart()
    On Error GoTo err
If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    End If
err:
    If err.Number <> 0 Then MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation
    
End Sub

Private Sub Detail_Format()
Dim Mov1 As Double, Mov2 As Double


 If Me.Field26.Text = "0.00" Or Me.Field26.Text = "" Then
   Mov1 = 0
 Else
   Mov1 = Me.Field26.Text
 End If
 
  If Me.Field27.Text = "0.00" Or Me.Field27.Text = "" Then
   Mov2 = 0
 Else
   Mov2 = Me.Field27.Text
 End If
      TipoCuenta = Me.FldTipoCuenta.Text

If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Cuentas x Pagar" Or TipoCuenta = "Cuentas de Gastos" Or TipoCuenta = "Bancos" Then
 SaldoFin = Mov1 - Mov2 + SaldoFin
    Me.FldSaldo.Text = Format(SaldoIni + SaldoFin, "###,###,###,##0.#0")
Else
 SaldoFin = Mov2 - Mov1 + SaldoFin
    Me.FldSaldo.Text = Format(SaldoIni - SaldoFin, "###,###,###,##0.#0")
End If

'pone al tipo sólo D o C
Me.Field22.Text = Mid(Me.Field22, 1, 1)
End Sub

Private Sub GroupFooter2_Format()
'///////////////Busco el Acumulado de la cuenta hasta la ultima fecha Seleccionada////////////
'///////////////////////////////////////////////////////////////////////////////////////////////
    NumFecha1 = FrmReportes.DTFecha1.Value
    NumFecha2 = FrmReportes.DTFecha2.Value
    If FrmReportes.DBCodigo.Text = "" Then
     CodigoCuenta = Me.Field16.Text
    Else
     CodigoCuenta = FrmReportes.DBCodigo.Text
    End If
'        FrmReportes.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito From Transacciones Where (((Transacciones.FechaTransaccion) <= " & NumFecha2 & ")) GROUP BY Transacciones.CodCuentas Having (((Transacciones.CodCuentas) = '" & CodigoCuenta & "'))"
        FrmReportes.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito From Transacciones Where (((Transacciones.FechaTransaccion) <= '" & Format(FrmReportes.DTFecha2, "yyyymmdd") & "')) GROUP BY Transacciones.CodCuentas Having (((Transacciones.CodCuentas) = '" & CodigoCuenta & "'))"
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

End Sub

Private Sub GroupHeader2_Format()
'///////////////Busco el Acumulado de la cuenta hasta la ultima fecha Seleccionada////////////
'///////////////////////////////////////////////////////////////////////////////////////////////
    SaldoFin = 0
    Debito = 0
    Credito = 0
    
    NumFecha1 = FrmReportes.DTFecha1.Value
    NumFecha2 = FrmReportes.DTFecha2.Value
    If FrmReportes.DBCodigo.Text = "" Then
     CodigoCuenta = Me.Field16.Text
    Else
     CodigoCuenta = FrmReportes.DBCodigo.Text
    End If
        FrmReportes.DtaHistorial.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Debito*TCambio) AS MDebito, Sum(TCambio*Credito) AS MCredito From Transacciones Where (((Transacciones.FechaTransaccion) < '" & Format(FrmReportes.DTFecha1, "yyyymmdd") & "')) GROUP BY Transacciones.CodCuentas Having (((Transacciones.CodCuentas) = '" & CodigoCuenta & "'))"
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
