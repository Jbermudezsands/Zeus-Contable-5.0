VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepTransacciones 
   Caption         =   "Reporte de Transacciones"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19368
   SectionData     =   "ArepTransacciones.dsx":0000
End
Attribute VB_Name = "ArepTransacciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FacturaNo As String, FechaFactura As Date, Tipo As String
Public Moneda As String
Private Sub ActiveReport_FetchData(EOF As Boolean)
    If Not EOF Then
    'Gets the current records SupplierID
      If Not IsNull(Me.DataControl1.Recordset.Fields("FacturaNo")) Then
        FacturaNo = Me.DataControl1.Recordset.Fields("FacturaNo")
      End If
      FechaFactura = Me.DataControl1.Recordset.Fields("FechaTransaccion")
      Tipo = Me.DataControl1.Recordset.Fields("Fuente")
      
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

QuienReporte = Me.Name

'    Me.LblEmpresa = MDIPrimero.AdoConfiguracion.Recordset("NombreEmpresa")
'    ArepTransacciones.LblEmpresa1 = MDIPrimero.AdoConfiguracion.Recordset("Direccion")
'    ArepTransacciones.LblEmpresa2 = "RUC: " & MDIPrimero.AdoConfiguracion.Recordset("NumeroRuc")
''    ArepTransacciones.Logo.Picture = LoadPicture(RutaLogo)
'    ArepTransacciones.LblFecha.Caption = Format(Now, "long")
    
             Me.LblEmpresa = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
             Me.LblEmpresa1 = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
             Me.LblEmpresa2 = "RUC: " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
             Me.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
'             Me.LblMoneda.Caption = "Expresado en " & Moneda
             
'             Me.LblCodigo.Caption = "Mis Code:" & FrmReportes.DBCodigo.Text
'             Me.LblRango.Caption = "Filtrado Desde: " & FrmReportes.CodDesde & " Hasta " & FrmReportes.CodHasta

  
If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    End If
err:
    If err.Number <> 0 Then MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation


End Sub

Private Sub Detail_Format()
If Me.Field11.Text = "0.00" Then
  Me.Field11.Visible = False
Else
  Me.Field11.Visible = True
End If

If Me.Field12.Text = "0.00" Then
  Me.Field12.Visible = False
Else
  Me.Field12.Visible = True
End If

Me.Field8.Hyperlink = Format(FechaFactura, "dd/mm/yyyy") & ";" & FacturaNo & ";" & Tipo

End Sub

Private Sub GroupHeader1_Format()
  Dim FechaTransaccion As String, NumeroTransaccion As Double
  
  If Me.Field1.Text <> "" Then
      FechaTransaccion = Format(CDate(Me.Field1.Text), "yyyy-mm-dd")
      NumeroTransaccion = Me.Field2.Text
    
    
      MDIPrimero.AdoConsulta.RecordSource = "SELECT * From IndiceTransaccion WHERE (FechaTransaccion = CONVERT(DATETIME, '" & FechaTransaccion & "', 102)) AND (NumeroMovimiento = " & NumeroTransaccion & ")"
      MDIPrimero.AdoConsulta.Refresh
      If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
        Me.LblMoneda.Caption = MDIPrimero.AdoConsulta.Recordset("TipoMoneda")
      
      End If
  
End If

End Sub

Private Sub ReportFooter_Format()
 Me.LblUsuario.Caption = NombreUsuario
End Sub
