VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepSuspense 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "ArepSuspense.dsx":0000
End
Attribute VB_Name = "ArepSuspense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
Dim CodCuentas As String, MontoI As Double, MontoJ As Double, MontoK As Double, Fechas As Date
'/////////Busco los moviminetos antes de la fecha///////////////////////
Fechas = Me.Field2.Text
NumFecha1 = Fechas
FrmReportes.DtaConsulta.RecordSource = "SELECT Sum(AdelantosJustifica.MontoAnticipo) AS SumaDeMontoAnticipo FROM Cuentas INNER JOIN AdelantosJustifica ON Cuentas.CodCuentas = AdelantosJustifica.CodCuenta Where (((Cuentas.DescripcionCuentas) = " & CodCuentas & ") And ((AdelantosJustifica.FechaAnticipo) <= " & NumFecha1 & "))"
FrmReportes.Refresh
 If Not FrmReportes.DtaConsulta.Recordset.EOF Then
  MontoI = FrmReportes.DtaConsulta.Recordset("SumaDeMontoAnticipo")
  Me.LblI.Caption = Format(MontoI, "##,##0.00")
 Else
    Me.LblI.Caption = Format(0, "##,##0.00")
 End If
 '///////////////Busco los movimientos del Periodo//////////////////////////
 FrmReportes.DtaConsulta.RecordSource = "SELECT Sum(AdelantosJustifica.MontoAnticipo) AS SumaDeMontoAnticipo FROM Cuentas INNER JOIN AdelantosJustifica ON Cuentas.CodCuentas = AdelantosJustifica.CodCuenta Where (((Cuentas.DescripcionCuentas) = " & CodCuentas & ") And ((AdelantosJustifica.FechaAnticipo) = " & NumFecha1 & "))"
 FrmReportes.Refresh
  If Not FrmReportes.DtaConsulta.Recordset.EOF Then
  MontoJ = FrmReportes.DtaConsulta.Recordset("SumaDeMontoAnticipo")
  Me.LblJ.Caption = Format(MontoJ, "##,##0.00")
 Else
    Me.LblJ.Caption = Format(0, "##,##0.00")
 End If



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

