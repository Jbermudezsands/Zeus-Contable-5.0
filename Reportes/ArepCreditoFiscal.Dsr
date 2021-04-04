VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepCreditoProveedor 
   Caption         =   "Reporte Anexo al Credito Fiscal Declaracion IVA"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   33655
   _ExtentY        =   20267
   SectionData     =   "ArepCreditoFiscal.dsx":0000
End
Attribute VB_Name = "ArepCreditoProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public GranTotal As Double, GranSubTotal As Double, GranIva As Double

Private Sub Detail_Format()
Dim CodCuenta As String, Factura As String, MontoIva As Double, Total As Double, SubTotal As Double

'///////////////////////////////////////////////////////////////////////////////////////////////
'///////////////BUSCO LOS DATOS DEL PROVEEDOR///////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////
SubTotal = 0
CodCuenta = Me.FldCodCuenta.Text
FrmReportes.DtaConsulta.RecordSource = "SELECT * From Cuentas WHERE (CodCuentas = '" & CodCuenta & "')"
FrmReportes.DtaConsulta.Refresh
If Not FrmReportes.DtaConsulta.Recordset.EOF Then
  Me.LblNombres.Caption = FrmReportes.DtaConsulta.Recordset("DescripcionCuentas")
  If Not IsNull(FrmReportes.DtaConsulta.Recordset("RUC")) Then
   Me.LblRUCS.Caption = FrmReportes.DtaConsulta.Recordset("RUC")
  End If
  If Not IsNull(FrmReportes.DtaConsulta.Recordset("DescRetencion")) Then
    Me.LblTasa.Caption = FrmReportes.DtaConsulta.Recordset("DescRetencion") & "%"
  End If
  
  If Not Me.FldFactura.Text = "" Then
    Factura = FldFactura.Text
  End If

End If

' If Me.FldSubTotal.Text <> "" Then
'  SubTotal = Me.FldSubTotal.Text
' End If

'/////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////BUSCO SI LA FACTURA FUE POR COMPRA DE SERVICIOS/////////////

FrmReportes.DtaConsulta.RecordSource = "SELECT *, Transacciones.TCambio * Transacciones.Debito + Transacciones.TCambio * Transacciones.Credito AS Monto, Cuentas.TipoCuenta FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas  " & _
                                       "WHERE     (Transacciones.FacturaNo = '" & Factura & "') AND (Cuentas.CausaIva = 0) AND (Cuentas.TipoCuenta = 'Gastos')"
FrmReportes.DtaConsulta.Refresh
Do While Not FrmReportes.DtaConsulta.Recordset.EOF
 SubTotal = FrmReportes.DtaConsulta.Recordset("Debito")
 FrmReportes.DtaConsulta.Recordset.MoveNext
Loop


'/////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////BUSCO SI LA FACTURA FUE POR COMPRA DE PRODUCTOS/////////////

FrmReportes.DtaConsulta.RecordSource = "SELECT *, Transacciones.TCambio * Transacciones.Debito + Transacciones.TCambio * Transacciones.Credito AS Monto, Cuentas.TipoCuenta FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas  " & _
                                       "WHERE     (Transacciones.FacturaNo = '" & Factura & "') AND (Cuentas.CausaIva = 0) AND (Cuentas.TipoCuenta = 'Inventario')"
FrmReportes.DtaConsulta.Refresh
Do While Not FrmReportes.DtaConsulta.Recordset.EOF
 SubTotal = FrmReportes.DtaConsulta.Recordset("Debito")
 FrmReportes.DtaConsulta.Recordset.MoveNext
Loop


'/////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////BUSCO SI LA FACTURA FUE POR COMPRA DE ACTIVO FIJO/////////////

FrmReportes.DtaConsulta.RecordSource = "SELECT *, Transacciones.TCambio * Transacciones.Debito + Transacciones.TCambio * Transacciones.Credito AS Monto, Cuentas.TipoCuenta FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas  " & _
                                       "WHERE     (Transacciones.FacturaNo = '" & Factura & "') AND (Cuentas.CausaIva = 0) AND (Cuentas.TipoCuenta = 'Activo Fijo')"
FrmReportes.DtaConsulta.Refresh
Do While Not FrmReportes.DtaConsulta.Recordset.EOF
 SubTotal = FrmReportes.DtaConsulta.Recordset("Debito")
 FrmReportes.DtaConsulta.Recordset.MoveNext
Loop

'/////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////BUSCO SI LA FACTURA FUE POR COMPRA DE ACTIVO FIJO/////////////

FrmReportes.DtaConsulta.RecordSource = "SELECT *, Transacciones.TCambio * Transacciones.Debito + Transacciones.TCambio * Transacciones.Credito AS Monto, Cuentas.TipoCuenta FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas  " & _
                                       "WHERE     (Transacciones.FacturaNo = '" & Factura & "') AND (Cuentas.CausaIva = 0) AND (Cuentas.TipoCuenta = 'Costos')"
FrmReportes.DtaConsulta.Refresh
Do While Not FrmReportes.DtaConsulta.Recordset.EOF
 SubTotal = FrmReportes.DtaConsulta.Recordset("Debito")
 FrmReportes.DtaConsulta.Recordset.MoveNext
Loop


 '///////////////////////////////////////////////////////////////////////////////////////////////////
 '/////////////BUSCO LA CUENTA DEL IMPUESTO IVA////////////////////////////////////////////////////////
 '/////////////////////////////////////////////////////////////////////////////////////////////////////
 FrmReportes.DtaConsulta2.RecordSource = "SELECT *, Transacciones.TCambio * Transacciones.Debito + Transacciones.TCambio * Transacciones.Credito AS Monto FROM Transacciones INNER JOIN  Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas WHERE (Cuentas.CausaIva = 1) AND (Transacciones.FacturaNo = '" & Factura & "')"
 FrmReportes.DtaConsulta2.Refresh
 If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
   Me.LblIVA.Caption = Format(FrmReportes.DtaConsulta2.Recordset("Monto"), "##,##0.00")
   MontoIva = FrmReportes.DtaConsulta2.Recordset("Monto")
   CodCuenta = FrmReportes.DtaConsulta2.Recordset("CodCuentas")
   FrmReportes.DtaConsulta.RecordSource = "SELECT * From Cuentas WHERE (CodCuentas = '" & CodCuenta & "')"
   FrmReportes.DtaConsulta.Refresh
   If Not FrmReportes.DtaConsulta.Recordset.EOF Then
     Me.LblTasa.Caption = FrmReportes.DtaConsulta.Recordset("DescRetencion")
 
   End If
   
   
 End If
 
 
 
 Total = SubTotal + MontoIva
 LblTotal.Caption = Format(Total, "##,##0.00")
 Me.LblSubTotal.Caption = Format(SubTotal, "##,##0.00")
 
 GranSubTotal = GranSubTotal + SubTotal
 GranIva = GranIva + MontoIva
 GranTotal = GranTotal + Total
 
End Sub

Private Sub ReportFooter_Format()
Me.LblGranIva.Caption = Format(GranIva, "##,##0.00")
Me.LblGranTotal.Caption = Format(GranTotal, "##,##0.00")
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

