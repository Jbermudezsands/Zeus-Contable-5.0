VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepLibroMayorSNiveles 
   Caption         =   "Reporte del Libro Diario Mayor"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19368
   SectionData     =   "ArepLibroMayorSNiveles.dsx":0000
End
Attribute VB_Name = "ArepLibroMayorSNiveles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalSaldoInicial As Double, TotalSaldoFinal As Double


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
Dim Tipo As String, Debe1 As Double, Haber1 As Double, Debe3 As Double, Haber3 As Double
Dim SaldoInicial As Double, SaldoFinal As Double

If Me.fldDebe1.Text <> "" Then
 Debe1 = Val(Me.fldDebe1.Text)
Else
 Debe1 = 0
End If

If Me.FldHaber1.Text <> "" Then
 Haber1 = Val(Me.FldHaber1.Text)
Else
 Haber1 = 0
End If

If Me.fldDebe3.Text <> "" Then
  Debe3 = Val(Me.fldDebe3.Text)
Else
  Debe3 = 0
End If

If Me.FldHaber3.Text <> "" Then
 Haber3 = Val(Me.FldHaber3.Text)
Else
 Haber3 = 0
End If




Tipo = Mid(Me.FldKeyGrupo.Text, 1, 1)



  If Tipo = "A" Or Tipo = "G" Or Tipo = "O" Then
     SaldoInicial = Debe1 - Haber1
     SaldoFinal = Debe3 - Haber3
     Me.LblSaldoFinal.Caption = Format(SaldoFinal, "##,##0.00")
     Me.LblSaldoInicial.Caption = Format(SaldoInicial, "##,##0.00")
  Else
     SaldoInicial = Haber1 - Debe1
     SaldoFinal = Haber3 - Debe3
     Me.LblSaldoFinal.Caption = Format(SaldoFinal, "##,##0.00")
     Me.LblSaldoInicial.Caption = Format(SaldoInicial, "##,##0.00")
  End If

    TotalSaldoInicial = SaldoInicial + TotalSaldoInicial
    TotalSaldoFinal = SaldoFinal + TotalSaldoFinal

End Sub

Private Sub ActiveReport_ReportStart()
    On Error GoTo err
    
    TotalSaldoInicial = 0
    TotalSaldoFinal = 0
    
If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    End If
err:
    If err.Number <> 0 Then MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation
    
End Sub

Private Sub GroupFooter1_Format()
  Me.LblTotalSaldoInicial.Caption = Format(TotalSaldoInicial, "##,##0.00")
  Me.LblTotalSaldoFinal.Caption = Format(TotalSaldoFinal, "##,##0.00")

End Sub
