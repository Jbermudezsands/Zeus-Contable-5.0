VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepBalanzaColumna 
   Caption         =   "Balanza de Comprobacion"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   19080
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   33655
   _ExtentY        =   19420
   SectionData     =   "ArepBalanzaColumna.dsx":0000
End
Attribute VB_Name = "ArepBalanzaColumna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalDebe3 As Double
Dim TotalHaber3 As Double

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
If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    End If
err:
    Exit Sub
End Sub

Private Sub Detail_Format()
Dim Nivel As Double
Dim TipoCuenta As String, Debito As Double, Credito As Double

'sólo ver el debe o el haber resultado según sea la naturaleza de la cuenta
'MsgBox Mid(Me.FldCuenta.Text, 1, 1)
If Me.fldDebe3.Text = "" Then Me.fldDebe3.Text = "0.00"
If Me.FldHaber3.Text = "" Then Me.FldHaber3.Text = "0.00"

If Mid(Me.FldCuenta.Text, 1, 1) = "1" Or Mid(Me.FldCuenta.Text, 1, 1) = "5" Then
   If CDbl(Me.fldDebe3.Text) > CDbl(Me.FldHaber3.Text) Then
    Me.fldDebe3.Text = CDbl(Me.fldDebe3.Text) - CDbl(Me.FldHaber3.Text)
    Me.FldHaber3.Text = "0.00"
   Else
    Me.FldHaber3.Text = CDbl(Me.FldHaber3.Text) - CDbl(Me.fldDebe3.Text)
    Me.fldDebe3.Text = "0.00"
   End If
ElseIf Mid(Me.FldCuenta.Text, 1, 1) = "2" Or Mid(Me.FldCuenta.Text, 1, 1) = "3" Or Mid(Me.FldCuenta.Text, 1, 1) = "4" Then
  If CDbl(Me.FldHaber3.Text) > CDbl(Me.fldDebe3.Text) Then
    Me.FldHaber3.Text = CDbl(Me.FldHaber3.Text) - CDbl(Me.fldDebe3.Text)
    Me.fldDebe3.Text = "0.00"
  Else
    Me.fldDebe3.Text = CDbl(Me.fldDebe3.Text) - CDbl(Me.FldHaber3.Text)
    Me.FldHaber3.Text = "0.00"
  End If
End If
TotalDebe3 = TotalDebe3 + CDbl(Me.fldDebe3.Text)
TotalHaber3 = TotalHaber3 + CDbl(Me.FldHaber3.Text)
Me.fldDebe3.Text = Format(Me.fldDebe3, "###,###,###,##0.00")
Me.FldHaber3.Text = Format(Me.FldHaber3, "###,###,###,##0.00")



 If Me.FldNivel.Text <> "" Then
  Nivel = Me.FldNivel.Text
 Else
  Nivel = 0
 End If
 


 


        TipoCuenta = Me.FldUbicacion.Text
        If Me.fldDebe1.Text = "" Then
         Debito = 0
        Else
         Debito = Me.fldDebe1.Text
        End If

        If Me.FldHaber1.Text = "" Then
         Credito = 0
        Else
         Credito = Me.FldHaber1.Text
        End If

        If TipoCuenta = "Cajas" Or TipoCuenta = "Bancos" Or TipoCuenta = "Inventario" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Terreno y Edificios" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Mobiliario y Equipo de Oficina" Or TipoCuenta = "Equipo Rodante" Or TipoCuenta = "Depreciacion Acumulada" Or TipoCuenta = "Papeleria y Utiles de Oficina" Or TipoCuenta = "Pagos Anticipados" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Compras" Or TipoCuenta = "Costos Produccion" Or TipoCuenta = "Costos Generales Produccion" Or TipoCuenta = "Acarreo y Fletes" Or TipoCuenta = "Rebajas y Dev S/Compra" Or TipoCuenta = "Sueldos y Comisiones" Or TipoCuenta = "Propaganda" Or TipoCuenta = "Sueldos Admon" Or TipoCuenta = "Energia y Agua Potable" Or TipoCuenta = "Comisiones/Intereses Gandados" Or TipoCuenta = "Comisiones/Intereses Pagados" Or TipoCuenta = "Otros Ingresos" Or TipoCuenta = "Otros Gastos" Or TipoCuenta = "Impuestos Pagados" Then
           Me.LblSaldoInicial.Caption = Format(Debito - Credito, "##,##0.00")
        Else
           Me.LblSaldoInicial.Caption = Format(Credito - Debito, "##,##0.00")
        End If
 
 
 
 
 Select Case Nivel
         Case 3
             Me.FldCuenta.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 9pt; "
        
 Case Else
        Me.FldCuenta.Style = "color: rgb(0,0,0); font-weight: (null); font-size: 9pt"
   
 End Select

End Sub

Private Sub GroupFooter1_Format()
   
   FrmReportes.DtaConsulta.RecordSource = "SELECT SUM(Debe1) AS Debe1, SUM(Haber1) AS Haber1, SUM(Debe2) AS Debe2, SUM(Haber2) AS Haber2, SUM(Debe3) AS Debe3, SUM(Haber3) AS Haber3 From Reportes ORDER BY MAX(Orden)"
   FrmReportes.DtaConsulta.Refresh
   If Not FrmReportes.DtaConsulta.Recordset.EOF Then
       Me.FldTDebe3.Text = Format(FrmReportes.DtaConsulta.Recordset("Debe3"), "##,##0.00")
       Me.FldTHaber3.Text = Format(FrmReportes.DtaConsulta.Recordset("Haber3"), "##,##0.00")
   End If
   
End Sub

