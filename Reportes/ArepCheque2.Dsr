VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepCheque2 
   Caption         =   "Reporte de Cheques"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   20108
   _ExtentY        =   19368
   SectionData     =   "ArepCheque2.dsx":0000
End
Attribute VB_Name = "ArepCheque2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalDebito As Double, TotalCredito As Double, TotalDebitoD As Double, TotalCreditoD As Double, Memo As String, Moneda As String, ChequeNo As String



Private Sub ActiveReport_ReportStart()
    On Error GoTo err
If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    End If
err:
    If err.Number <> 0 Then MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation
    
    
   Me.LblEmpresa = MDIPrimero.AdoConfiguracion.Recordset("NombreEmpresa")
   Me.LblEmpresa1 = MDIPrimero.AdoConfiguracion.Recordset("Direccion")
   Me.LblEmpresa2 = "RUC: " & MDIPrimero.AdoConfiguracion.Recordset("NumeroRuc")
'   Me.Label17.Caption = NombreUsuario
'
If Memo <> "" Then
 Me.LblMemo.Caption = Memo
End If
 
Me.LblMoneda.Caption = "Expresado en: " & Moneda

Me.LblChequeNo.Caption = ChequeNo
 
If Val(FrmReportes.TxtTransaccion.Text) <> 0 Then
       FrmReportes.DtaConsulta.RecordSource = "SELECT Transacciones.ChequeNo, Transacciones.FechaTransaccion, Transacciones.CodCuentas, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, TCambio*Debito AS MDebito, TCambio*Credito AS MCredito, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito, Transacciones.VoucherNo, Transacciones.ChequeNo, Beneficiario From Transacciones WHERE (((Transacciones.FechaTransaccion) Between '" & Format(FrmReportes.DTFecha1, "yyyymmdd") & "' And '" & Format(FrmReportes.DTFecha2, "yyyymmdd") & "') AND ((Transacciones.NumeroMovimiento)=" & Val(FrmReportes.TxtTransaccion.Text) & ") AND ((Transacciones.ChequeNo) Is Not Null))"
       FrmReportes.DtaConsulta.Refresh
       If Not FrmReportes.DtaConsulta.Recordset.EOF Then
        If Not IsNull(FrmReportes.DtaConsulta.Recordset("ChequeNo")) Then
            If Not IsNumeric(FrmReportes.DtaConsulta.Recordset("ChequeNo")) Then
'                MsgBox "Cheque con No inválido: " & FrmReportes.DtaConsulta.Recordset("ChequeNo"), vbInformation
            Else
                Me.LblChequeNo.Caption = FrmReportes.DtaConsulta.Recordset("ChequeNo")
                Me.LblNombre.Caption = FrmReportes.DtaConsulta.Recordset("Beneficiario")
            End If
        End If
       End If
 End If
End Sub

Private Sub Detail_Format()
Dim Debito As Double, Credito As Double, DebitoD As Double, CreditoD As Double

Debito = Me.FldDebito.Text
Credito = Me.FldCredito.Text
DebitoD = Me.FldDebitoD.Text
CreditoD = Me.FldCreditoD.Text

TotalDebito = TotalDebito + Debito
TotalCredito = TotalCredito + Credito

TotalDebitoD = TotalDebitoD + DebitoD
TotalCreditoD = TotalCreditoD + CreditoD


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

Private Sub PageFooter_Format()

     
   MDIPrimero.AdoConsulta.RecordSource = "SELECT MAX(Transacciones.CodCuentas) AS Expr1, Transacciones.NumeroMovimiento, SUM(Transacciones.Debito) AS Debito, SUM(Transacciones.Credito) AS Credito,  SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito / Tasas.MontoCordobas ELSE Transacciones.Debito END) AS DebitoD, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito / Tasas.MontoCordobas ELSE Transacciones.Credito END) AS CreditoD FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento  " & _
                                         "WHERE (Transacciones.FechaTransaccion = CONVERT(DATETIME,  '" & FechaCheque & "', 102)) GROUP BY Transacciones.NumeroMovimiento Having (Transacciones.NumeroMovimiento = " & NumeroMovimientos & ")"
   MDIPrimero.AdoConsulta.Refresh
   If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
        
        Me.LblTotalDebito.Caption = Format(MDIPrimero.AdoConsulta.Recordset("Debito"), "##,##0.00")
        Me.LblTotalCredito.Caption = Format(MDIPrimero.AdoConsulta.Recordset("Credito"), "##,##0.00")
        
        Me.LblTotalDebitoD.Caption = Format(MDIPrimero.AdoConsulta.Recordset("DebitoD"), "##,##0.00")
        Me.LblTotalCreditoD.Caption = Format(MDIPrimero.AdoConsulta.Recordset("CreditoD"), "##,##0.00")
   
   End If
     

' Me.Label17.Caption = NombreUsuario
 
     
'Me.LblTotalDebito.Caption = Format(TotalDebito, "##,##0.00")
'Me.LblTotalCredito.Caption = Format(TotalCredito, "##,##0.00")
'
'Me.LblTotalDebitoD.Caption = Format(TotalDebitoD, "##,##0.00")
'Me.LblTotalCreditoD.Caption = Format(TotalCreditoD, "##,##0.00")
    
End Sub

Private Sub PageHeader_Format()
Dim TasaCambio As Double, Fecha As Date
TotalCredito = 0
TotalDebito = 0
TotalCreditoD = 0
TotalDebitoD = 0

Fecha = Me.Field13.Text
TasaCambio = BuscaTasaCambio(Fecha)
Me.LblTasaCambio.Caption = "Tasa del Dia: " & Format(TasaCambio, "##,##0.0000")


End Sub
