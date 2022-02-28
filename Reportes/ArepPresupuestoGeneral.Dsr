VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepPresupuestoGeneral 
   Caption         =   "Presupuesto General"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35930
   _ExtentY        =   19420
   SectionData     =   "ArepPresupuestoGeneral.dsx":0000
End
Attribute VB_Name = "ArepPresupuestoGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
On Error GoTo err
 
     Me.Label13.Caption = "PRESUPUESTO PARA EL AÑO " & Year(Fecha1)
     Me.LblEmpresa = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
     Me.LblEmpresa1 = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
     Me.LblEmpresa2 = "RUC: " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
     Me.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
     Me.Logo.Picture = LoadPicture(RutaLogo)

 
If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    End If
err:
    If err.Number <> 0 Then MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation
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


Private Sub Detail_Format()
 Dim Acumulado As Double

 Acumulado = Int(Me.FldEnero.Text) + Int(Me.FldMarzo.Text) + Int(Me.FldFebrero.Text) + Int(Me.FldAbril.Text) + Int(Me.FldMayo.Text) + Int(Me.FldJunio.Text) + Int(Me.FldJulio.Text) + Int(Me.FldAgosto.Text) + Int(Me.FldSeptiembre.Text) + Int(Me.FldOctubre.Text) + Int(Me.FldOctubre.Text) + Int(Me.FldNoviembre.Text) + Int(Me.FldDiciembre.Text)

End Sub
