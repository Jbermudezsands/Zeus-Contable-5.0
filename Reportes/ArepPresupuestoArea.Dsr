VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepPresupuestoArea 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35930
   _ExtentY        =   19420
   SectionData     =   "ArepPresupuestoArea.dsx":0000
End
Attribute VB_Name = "ArepPresupuestoArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FechaIni As String, FechaFin As String


Private Sub ActiveReport_ReportStart()
 On Error GoTo err
 
     Me.Label13.Caption = "PRESUPUESTO DESDE EL " & FechaIni & " HASTA " & FechaFin
    
     Me.LblEmpresa.Caption = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
     Me.LblEmpresa1.Caption = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
     Me.LblEmpresa2.Caption = "RUC: " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
     Me.LblFechaImpreso.Caption = Format(Now, "dd/mm/yyyy")
      Me.Logo.Picture = LoadPicture(RutaLogo)
 
If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
 End If
err:
    If err.Number <> 0 Then MsgBox "Hay un problema con la direcci�n del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuraci�n del Sistema", vbInformation
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
    ' Presentar el cuadro de di�logo Abrir
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

