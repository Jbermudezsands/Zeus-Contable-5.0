VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepCheque 
   Caption         =   "Reporte de Cheques"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20340
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35878
   _ExtentY        =   19420
   SectionData     =   "ArepCheque.dsx":0000
End
Attribute VB_Name = "ArepCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalDebito As Double, TotalCredito As Double



Private Sub ActiveReport_ReportStart()
    On Error GoTo err
If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    End If
err:
    If err.Number <> 0 Then MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation
    
    
   Me.LblEmpresa = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
   Me.LblEmpresa1 = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
   Me.LblEmpresa2 = "RUC: " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
   Me.Label17.Caption = NombreUsuario

If FrmCheque.TxtMemo.Text <> "" Then
 Me.LblMemo.Caption = FrmCheque.TxtMemo.Text
End If
 
Me.LblMoneda.Caption = "Expresado en: " & FrmCheque.CmbMoneda.Text
 
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
Dim Debito As Double, Credito As Double

Debito = Me.FldDebito.Text
Credito = Me.FldCredito.Text

TotalDebito = TotalDebito + Debito
TotalCredito = TotalCredito + Credito


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
 Me.Label17.Caption = NombreUsuario
 
     
Me.LblTotalDebito.Caption = Format(TotalDebito, "##,##0.00")
Me.LblTotalCredito.Caption = Format(TotalCredito, "##,##0.00")
    
End Sub

Private Sub PageHeader_Format()
Dim TasaCambio As Double, Fecha As Date
TotalCredito = 0
TotalDebito = 0

Fecha = Me.Field13.Text
TasaCambio = BuscaTasaCambio(Fecha)
Me.LblTasaCambio.Caption = "Tasa del Dia: " & Format(TasaCambio, "##,##0.0000")


End Sub
