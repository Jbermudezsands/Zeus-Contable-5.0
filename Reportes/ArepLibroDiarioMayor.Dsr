VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepLibroDiarioMayor 
   Caption         =   "Reporte del Libro Diario Mayor"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19368
   SectionData     =   "ArepLibroDiarioMayor.dsx":0000
End
Attribute VB_Name = "ArepLibroDiarioMayor"
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

Private Sub Detail_Format()
Dim NumeroFecha As Long, Fecha As Date
Dim llaveGrupo As String, SaldoInicial As Double
Dim SaldoFinal As Double, Tipo As String, Debito As Double
Dim Credito As Double

If Me.FldKeyGrupo.Text = "B01000001000000" Then
 Cod = 1
End If
llaveGrupo = Me.FldKeyGrupo.Text
If Not Me.FldFechaTransaccion.Text = "" Then
 Fecha = Me.FldFechaTransaccion.Text
End If
NumeroFecha = Fecha
Tipo = Mid(llaveGrupo, 1, 1)

SaldoInicial = 0
SaldoFinal = 0
Debito = 0
Credito = 0


SQL = "SELECT MAX(Transacciones.CodCuentas) AS ÚltimoDeCodCuentas, MAX(Transacciones.NumeroMovimiento) AS ÚltimoDeNumeroMovimiento, MAX(Transacciones.NombreCuenta) AS ÚltimoDeNombreCuenta, Avg(Transacciones.TCambio) AS PromedioDeTCambio, Sum(Transacciones.TCambio*Transacciones.Debito) AS Debito, Sum(Transacciones.TCambio*Transacciones.Credito) AS Credito, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, Sum(Transacciones.TCambio*Transacciones.Debito)-Sum(Transacciones.TCambio*Transacciones.Credito) AS Saldo1, Sum(Transacciones.TCambio*Transacciones.Credito)-Sum(Transacciones.TCambio*Transacciones.Debito) AS Saldo2 " & _
"FROM (Grupos INNER JOIN Cuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo) INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas " & _
"Where (((Transacciones.FechaTransaccion) < '" & Format(Fecha, "YYYYMMDD") & "')) GROUP BY Cuentas.KeyGrupo, Cuentas.DescripcionGrupo " & _
"Having (((Cuentas.KeyGrupo) = '" & llaveGrupo & "')) ORDER BY Cuentas.KeyGrupo "

FrmReportes.DtaConsulta.RecordSource = SQL
FrmReportes.DtaConsulta.Refresh
If Not FrmReportes.DtaConsulta.Recordset.EOF Then
  If Tipo = "A" Or Tipo = "G" Or Tipo = "O" Then
     SaldoInicial = FrmReportes.DtaConsulta.Recordset("Saldo1")
     
'     Debito = Me.FldDebito.Text
'     Credito = Me.FldCredito.Text
     If Debito > Credito Then
      Me.LblDebito.Caption = Format(Debito - Credito, "##,##0.00")
      Me.LblCredito.Caption = ""
     Else
      Me.LblCredito.Caption = Format(Credito - Debito, "##,##0.00")
      Me.LblDebito.Caption = ""
     End If
     SaldoFinal = SaldoInicial + Debito - Credito
     Me.LblSaldoFinal.Caption = Format(SaldoFinal, "##,##0.00")
     Me.LblSaldoInicial.Caption = Format(SaldoInicial, "##,##0.00")
  Else
     SaldoInicial = FrmReportes.DtaConsulta.Recordset("Saldo2")
'     Debito = Me.FldDebito.Text
'     Credito = Me.FldCredito.Text
     If Credito > Debito Then
      Me.LblCredito.Caption = Format(Credito - Debito, "##,##0.00")
      Me.LblDebito.Caption = ""
     Else
      Me.LblDebito.Caption = Format(Debito - Credito, "##,##0.00")
      Me.LblCredito.Caption = ""
     End If
     SaldoFinal = SaldoInicial + Credito - Debito
     Me.LblSaldoFinal.Caption = Format(SaldoFinal, "##,##0.00")
     Me.LblSaldoInicial.Caption = Format(SaldoInicial, "##,##0.00")
  End If
Else
     If Tipo = "A" Or Tipo = "G" Or Tipo = "O" Then
     SaldoInicial = 0
     If Not Me.FldDebito.Text = "" Then
     Debito = Me.FldDebito.Text
     End If
     If Not Me.FldCredito.Text = "" Then
     Credito = Me.FldCredito.Text
     End If
     
     If Debito > Credito Then
      Me.LblDebito.Caption = Format(Debito - Credito, "##,##0.00")
      Me.LblCredito.Caption = ""
     Else
      Me.LblCredito.Caption = Format(Credito - Debito, "##,##0.00")
      Me.LblDebito.Caption = ""
     End If
     SaldoFinal = SaldoInicial + Debito - Credito
     Me.LblSaldoFinal.Caption = Format(SaldoFinal, "##,##0.00")
     Me.LblSaldoInicial.Caption = Format(SaldoInicial, "##,##0.00")
  Else
     SaldoInicial = 0
     If Not Me.FldDebito.Text = "" Then
     Debito = Me.FldDebito.Text
     End If
     If Not Me.FldCredito.Text = "" Then
     Credito = Me.FldCredito.Text
     End If
     
     If Credito > Debito Then
      Me.LblCredito.Caption = Format(Credito - Debito, "##,##0.00")
      Me.LblDebito.Caption = ""
     Else
      Me.LblDebito.Caption = Format(Debito - Credito, "##,##0.00")
      Me.LblCredito.Caption = ""
     End If
     SaldoFinal = SaldoInicial + Credito - Debito
     Me.LblSaldoFinal.Caption = Format(SaldoFinal, "##,##0.00")
     Me.LblSaldoInicial.Caption = Format(SaldoInicial, "##,##0.00")
  End If
End If

End Sub

Private Sub ActiveReport_ReportStart()
    On Error GoTo err
If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    End If
err:
    If err.Number <> 0 Then MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation
    
End Sub

