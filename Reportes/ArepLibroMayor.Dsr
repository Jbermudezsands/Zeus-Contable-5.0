VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepLibroMayor 
   Caption         =   "Reporte del Libro Diario Mayor"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20340
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35878
   _ExtentY        =   19420
   SectionData     =   "ArepLibroMayor.dsx":0000
End
Attribute VB_Name = "ArepLibroMayor"
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
'Dim Tipo As String, Debe1 As Double, Haber1 As Double, Debe3 As Double, Haber3 As Double
'Dim SaldoInicial As Double, SaldoFinal As Double
'
'If Me.FldDebe1.Text <> "" Then
' Debe1 = Val(Me.FldDebe1.Text)
'Else
' Debe1 = 0
'End If
'
'If Me.FldHaber1.Text <> "" Then
' Haber1 = Val(Me.FldHaber1.Text)
'Else
' Haber1 = 0
'End If
'
'If Me.FldDebe3.Text <> "" Then
'  Debe3 = Val(Me.FldDebe3.Text)
'Else
'  Debe3 = 0
'End If
'
'If Me.FldHaber3.Text <> "" Then
' Haber3 = Val(Me.FldHaber3.Text)
'Else
' Haber3 = 0
'End If
'
'
'
'
'Tipo = Mid(Me.FldKeyGrupo.Text, 1, 1)
'
'
'
'  If Tipo = "A" Or Tipo = "G" Or Tipo = "O" Then
'     SaldoInicial = Debe1 - Haber1
'     SaldoFinal = Debe3 - Haber3
'     Me.LblSaldoFinal.Caption = Format(SaldoFinal, "##,##0.00")
'     Me.LblSaldoInicial.Caption = Format(SaldoInicial, "##,##0.00")
'  Else
'     SaldoInicial = Haber1 - Debe1
'     SaldoFinal = Haber3 - Debe3
'     Me.LblSaldoFinal.Caption = Format(SaldoFinal, "##,##0.00")
'     Me.LblSaldoInicial.Caption = Format(SaldoInicial, "##,##0.00")
'  End If
'
'    TotalSaldoInicial = SaldoInicial + TotalSaldoInicial
'    TotalSaldoFinal = SaldoFinal + TotalSaldoFinal

    Dim SqlString As String, Registro As Double
    
    Debito = 0
    Credito = 0
    TotalInicial = 0
    TotalFinal = 0
    
   '-----------------------------------------------------------------------------------
   '--------------------------CARGO LOS ACTIVOS ---------------------------------------
   '-----------------------------------------------------------------------------------
    SqlString = "SELECT  * From Reportes WHERE (Nivel = '" & FrmReportes.CmbNivel2.Text & "') AND (Descripcion LIKE N'%Total%') AND (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 <> 0) AND (KeyGrupo LIKE N'%A%')ORDER BY Orden"
    Set Me.SrptActivos.object = New ArepSubReporteLibroMayor
    Me.SrptActivos.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptActivos.object.DataControl1.Source = SqlString
    
    
   '-----------------------------------------------------------------------------------
   '--------------------------CARGO LOS PASIVOS---------------------------------------
   '-----------------------------------------------------------------------------------
    SqlString = "SELECT  * From Reportes WHERE (Nivel = '" & FrmReportes.CmbNivel3.Text & "') AND (Descripcion LIKE N'%Total%') AND (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 <> 0) AND (KeyGrupo LIKE N'%B%')ORDER BY Orden"
    Set Me.SrptPasivos.object = New ArepSubReporteLibroMayor
    Me.SrptPasivos.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptPasivos.object.DataControl1.Source = SqlString
    
   '-----------------------------------------------------------------------------------
   '--------------------------CARGO EL CAPITAL---------------------------------------
   '-----------------------------------------------------------------------------------
    SqlString = "SELECT  * From Reportes WHERE (Nivel = '" & FrmReportes.CmbNivel4.Text & "') AND (Descripcion LIKE N'%Total%') AND (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 <> 0) AND (KeyGrupo LIKE N'%C%')ORDER BY Orden"
    Set Me.SrptCapital.object = New ArepSubReporteLibroMayor
    Me.SrptCapital.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptCapital.object.DataControl1.Source = SqlString
    
   '-----------------------------------------------------------------------------------
   '--------------------------CARGO EL INGRESO---------------------------------------
   '-----------------------------------------------------------------------------------
    SqlString = "SELECT  * From Reportes WHERE (Nivel = '" & FrmReportes.CmbNivel5.Text & "') AND (Descripcion LIKE N'%Total%') AND (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 <> 0) AND (KeyGrupo LIKE N'%D%')ORDER BY Orden"
    Set Me.SrptIngresos.object = New ArepSubReporteLibroMayor
    Me.SrptIngresos.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptIngresos.object.DataControl1.Source = SqlString
    
   '-----------------------------------------------------------------------------------
   '--------------------------CARGO LOS COSTOS---------------------------------------
   '-----------------------------------------------------------------------------------
    SqlString = "SELECT  * From Reportes WHERE (Nivel = '" & FrmReportes.CmbNivel6.Text & "') AND (Descripcion LIKE N'%Total%') AND (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 <> 0) AND (KeyGrupo LIKE N'%G%')ORDER BY Orden"
    Set Me.SrptCostos.object = New ArepSubReporteLibroMayor
    Me.SrptCostos.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptCostos.object.DataControl1.Source = SqlString
    
   '-----------------------------------------------------------------------------------
   '--------------------------CARGO LOS GASTOS---------------------------------------
   '-----------------------------------------------------------------------------------
    SqlString = "SELECT  * From Reportes WHERE (Nivel = '" & FrmReportes.CmbNivel7.Text & "') AND (Descripcion LIKE N'%Total%') AND (Debe1 + Haber1 + Debe2 + Haber2 + Debe3 + Haber3 <> 0) AND (KeyGrupo LIKE N'%O%')ORDER BY Orden"
    Set Me.SrptGastos.object = New ArepSubReporteLibroMayor
    Me.SrptGastos.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptGastos.object.DataControl1.Source = SqlString

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
  Me.LblTotalSaldoInicial.Caption = Format(TotalInicial, "##,##0.00")
  Me.LblTotalSaldoFinal.Caption = Format(TotalFinal, "##,##0.00")
  
  Me.LblTotalDebito.Caption = Format(Debito, "##,##0.00")
  Me.LblTotalCredito.Caption = Format(Credito, "##,##0.00")

End Sub

