VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepLibroDiario 
   Caption         =   "Reporte del Libro Diario Mayor"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepLibroDiario.dsx":0000
End
Attribute VB_Name = "ArepLibroDiario"
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
    If err.Number <> 0 Then MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation
    
End Sub

Private Sub Detail_Format()
    Dim SqlString As String, Registro As Double
    
    Debito = 0
    Credito = 0
    TotalInicial = 0
    TotalFinal = 0
    
   '-----------------------------------------------------------------------------------
   '--------------------------CARGO LOS ACTIVOS ---------------------------------------
   '-----------------------------------------------------------------------------------
    SqlString = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden, CodCuentas, Ubicacion From Reportes WHERE (Descripcion LIKE N'%Total%') AND (Debe2 + Haber2 <> 0) AND (KeyGrupo LIKE N'%A%') AND (Nivel = '" & FrmReportes.CmbNivel2.Text & "') ORDER BY Orden"
    Set Me.SrptActivos.object = New ArepSubReporteLibroDiario
    Me.SrptActivos.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptActivos.object.DataControl1.Source = SqlString
    
    
   '-----------------------------------------------------------------------------------
   '--------------------------CARGO LOS PASIVOS---------------------------------------
   '-----------------------------------------------------------------------------------
    SqlString = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden, CodCuentas, Ubicacion From Reportes WHERE (Descripcion LIKE N'%Total%') AND (Debe2 + Haber2 <> 0) AND (KeyGrupo LIKE N'%B%') AND (Nivel = '" & FrmReportes.CmbNivel3.Text & "') ORDER BY Orden"
    Set Me.SrptPasivos.object = New ArepSubReporteLibroDiario
    Me.SrptPasivos.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptPasivos.object.DataControl1.Source = SqlString
    
   '-----------------------------------------------------------------------------------
   '--------------------------CARGO EL CAPITAL---------------------------------------
   '-----------------------------------------------------------------------------------
    SqlString = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden, CodCuentas, Ubicacion From Reportes WHERE (Descripcion LIKE N'%Total%') AND (Debe2 + Haber2 <> 0) AND (KeyGrupo LIKE N'%C%') AND (Nivel = '" & FrmReportes.CmbNivel4.Text & "') ORDER BY Orden"
    Set Me.SrptCapital.object = New ArepSubReporteLibroDiario
    Me.SrptCapital.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptCapital.object.DataControl1.Source = SqlString
    
   '-----------------------------------------------------------------------------------
   '--------------------------CARGO EL INGRESO---------------------------------------
   '-----------------------------------------------------------------------------------
    SqlString = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden, CodCuentas, Ubicacion From Reportes WHERE (Descripcion LIKE N'%Total%') AND (Debe2 + Haber2 <> 0) AND (KeyGrupo LIKE N'%D%') AND (Nivel = '" & FrmReportes.CmbNivel5.Text & "') ORDER BY Orden"
    Set Me.SrptIngresos.object = New ArepSubReporteLibroDiario
    Me.SrptIngresos.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptIngresos.object.DataControl1.Source = SqlString
    
   '-----------------------------------------------------------------------------------
   '--------------------------CARGO LOS COSTOS---------------------------------------
   '-----------------------------------------------------------------------------------
    SqlString = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden, CodCuentas, Ubicacion From Reportes WHERE (Descripcion LIKE N'%Total%') AND (Debe2 + Haber2 <> 0) AND (KeyGrupo LIKE N'%G%') AND (Nivel = '" & FrmReportes.CmbNivel6.Text & "') ORDER BY Orden"
    Set Me.SrptCostos.object = New ArepSubReporteLibroDiario
    Me.SrptCostos.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptCostos.object.DataControl1.Source = SqlString
    
   '-----------------------------------------------------------------------------------
   '--------------------------CARGO LOS GASTOS---------------------------------------
   '-----------------------------------------------------------------------------------
    SqlString = "SELECT Descripcion, Debe1, Haber1, Debe2, Haber2, Debe3, Haber3, KeyGrupo, KeyGrupoSuperior, KeyGrupoCuenta, Nivel, Orden, CodCuentas, Ubicacion From Reportes WHERE (Descripcion LIKE N'%Total%') AND (Debe2 + Haber2 <> 0) AND (KeyGrupo LIKE N'%O%') AND (Nivel = '" & FrmReportes.CmbNivel7.Text & "') ORDER BY Orden"
    Set Me.SrptGastos.object = New ArepSubReporteLibroDiario
    Me.SrptGastos.object.DataControl1.ConnectionString = ConexionReporte
    Me.SrptGastos.object.DataControl1.Source = SqlString

End Sub

Private Sub GroupFooter1_Format()
  Me.LblTotalDebito.Caption = Format(Debito, "##,##0.00")
  Me.LblTotalCredito.Caption = Format(Credito, "##,##0.00")
End Sub
