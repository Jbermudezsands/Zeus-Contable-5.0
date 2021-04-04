VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepBakControlSub 
   ClientHeight    =   11490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   19080
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   33655
   _ExtentY        =   20267
   SectionData     =   "BakControlSub.dsx":0000
End
Attribute VB_Name = "ArepBakControlSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_BeforePrint()
Detail.ColumnSpacing = 1
If CodigoBanco = Me.Field4 Then

 Me.Field4.Visible = False
 Me.Field5.Visible = False
 Me.LblSaldo.Visible = False
  Me.Label24.Visible = False
  Me.Label25.Visible = False
 Me.Label26.Visible = False
 Me.Label27.Visible = False
 
Else
 Me.Field4.Visible = True
 Me.Field5.Visible = True
 Me.LblSaldo.Visible = True
 Me.Label24.Visible = True
 Me.Label25.Visible = True
 Me.Label26.Visible = True
 Me.Label27.Visible = True
End If
End Sub

Private Sub Detail_Format()
If Not Me.Field11.Text = "" Then
  Debito = Me.Field11
  Debitos = Debitos + Debito
Else
  Debito = 0
End If

If Not Me.Field12.Text = "" Then
 Credito = Me.Field12
 Creditos = Creditos + Credito
Else
 Credito = 0
End If
Me.LblSaldo.Caption = Format(Debito - Credito, "##,##0.00")

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

