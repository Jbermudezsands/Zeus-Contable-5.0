VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepBalanceHistorico 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20340
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35878
   _ExtentY        =   19420
   SectionData     =   "ArepBalanceHistorico.dsx":0000
End
Attribute VB_Name = "ArepBalanceHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportEnd()
 On Error GoTo err:
   Dim RutaArchivo As String


If FrmReportes.ChkExportar.Value = 1 Then
'    Establecer CancelError a True
    FrmReportes.CDRuta.CancelError = True
'     Establecer los indicadores
    FrmReportes.CDRuta.Flags = cdlOFNHideReadOnly
'     Establecer los filtros
    FrmReportes.CDRuta.Filter = "Excel (*.XLS)|*.xls"
'     Especificar el filtro predeterminado
    FrmReportes.CDRuta.FilterIndex = 2
    ' Presentar el cuadro de diálogo Abrir
    FrmReportes.CDRuta.ShowSave
'    ' Presentar el nombre del archivo seleccionado
    RutaArchivo = FrmReportes.CDRuta.FileName  'varible que le doy la ruta
   
    MousePointer = 11
    
    Dim myExportObject As ActiveReportsExcelExport.ARExportExcel
    Dim Nombre As String
    
    Set myExportObject = CreateObject("ActiveReportsExcelExport.ARExportExcel")
    myExportObject.FileName = RutaArchivo
    myExportObject.Export Me.Pages
    Set myExportObject = Nothing

End If
err:
    If err.Number <> 0 Then Exit Sub

End Sub

Private Sub ActiveReport_ReportStart()
 On Error GoTo err
 Dim NumeroPeriodo1 As Double, NumeroPerido2 As Double, NumeroTabla As Integer, i As Double, FechaIni As String, FechaFin As String
Dim NumFecha2 As String, NumFecha1 As String

 On Error GoTo err
If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    End If
'err:
'    If err.Number <> 0 Then MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation

    NumeroPeriodo1 = FrmReportes.CmbIni.Text
    NumeroPeriodo2 = FrmReportes.CmbFin.Text
    
    If FrmReportes.Option8 = True Then
     NumeroTabla = 1
    ElseIf FrmReportes.Option7 = True Then
      NumeroTabla = 2
    ElseIf FrmReportes.Option6 = True Then
      NumeroTabla = 3
    End If
    
      FrmReportes.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) Between " & NumeroPeriodo1 & " And " & NumeroPeriodo2 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
      FrmReportes.DtaConsulta.Refresh
       FrmReportes.DtaConsulta.Recordset.MoveLast
       i = FrmReportes.DtaConsulta.Recordset.RecordCount
       FrmReportes.DtaConsulta.Recordset.MoveFirst
      Do While Not FrmReportes.DtaConsulta.Recordset.EOF


        If i = 1 Then
          FechaIni = "01/" & Month(FrmReportes.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
          FechaFin = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
        Else

         If NumeroPeriodo1 = FrmReportes.DtaConsulta.Recordset("Periodo") Then
          FechaIni = "01/" & Month(FrmReportes.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(FrmReportes.DtaConsulta.Recordset("FechaPeriodo"))
          NumFecha1 = FechaIni
         ElseIf NumeroPeriodo2 = FrmReportes.DtaConsulta.Recordset("Periodo") Then
          FechaFin = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
          NumFecha2 = FrmReportes.DtaConsulta.Recordset("FechaPeriodo")
         End If
        End If
        FrmReportes.DtaConsulta.Recordset.MoveNext
      Loop


    Me.LblMoneda.Caption = "Expresado en " & FrmReportes.CmbMoneda.Text
    If Dir(RutaLogo) <> "" Then
     Me.Logo.Picture = LoadPicture(RutaLogo)
    End If
    Me.LblEmpresa = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
    Me.LblEmpresa1 = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
    Me.LblEmpresa2 = "RUC: " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
    Me.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
    Me.LblFechaFin = Format(FechaFin, "dd/mm/yyyy")
    
    Me.LblFechaIni = Format(FechaIni, "dd/mm/yyyy")
 
 
 
 
   If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    End If
err:
    If err.Number <> 0 Then MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation
    End Sub

Private Sub Detail_Format()
Dim Descripcion As String, Monto As Double
Descripcion = Me.Field1.Text
Monto = Val(Me.Field3.Text) + Val(Me.Field7.Text) + Val(Me.Field4.Text)
 If Me.FLdKeyGrupo.Text = "A" Or Me.FLdKeyGrupo.Text = "PC" Then
  If Val(Monto) <> 0 Then
   Me.Line1.Visible = True
   Me.Line2.Visible = True
  Else
   Me.Line1.Visible = False
   Me.Line2.Visible = False
  End If
 Else
  Me.Line1.Visible = False
  Me.Line2.Visible = False
 End If
 
 If Me.FldNivel.Text = "2" Then
  If Val(Monto) <> 0 Then
    Me.Line5.Visible = True
  Else
    Me.Line5.Visible = False
  End If
 Else
  Me.Line5.Visible = False
 End If
 
           If FrmReportes.CmbNivel.Text > Me.FldNivel.Text Then
                        Me.Field1.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 9pt; "
                        Me.Field2.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                        Me.Field3.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                        Me.Field4.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                        Me.Field5.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                        Me.Field6.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
                        Me.Field7.Style = "color: rgb(0,0,0); font-weight: bold; font-size: 8pt; "
          Else
                     Me.Field1.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 8pt"
                     Me.Field2.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 8pt; "
                     Me.Field3.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 8pt; "
                     Me.Field4.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 8pt; "
                     Me.Field5.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 8pt; "
                     Me.Field6.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 8pt; "
                     Me.Field7.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 8pt; "
          End If

End Sub

