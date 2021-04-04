VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepBalanzaMayor 
   Caption         =   "Reporte Balanza ordenada a Nivel Mayor"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "ArepBalanzaMayor.dsx":0000
End
Attribute VB_Name = "ArepBalanzaMayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalSaldoInicial As Double, TotalSaldoFinal As Double, TotalDebitoIni As Double, TotalCreditoIni As Double



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

Private Sub Detail_Format()
Dim NumeroFecha As Long, Fecha As String
Dim llaveGrupo As String, SaldoInicial As Double
Dim SaldoFinal As Double, Tipo As String, Debito As Double
Dim Credito As Double, Fecha2 As String

If Me.FldKeyGrupo.Text = "B01000001000000" Then
 cod = 1
End If
llaveGrupo = Me.FldKeyGrupo.Text
Fecha = Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd")
Fecha2 = Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd")
'NumeroFecha = Fecha
Tipo = Mid(llaveGrupo, 1, 1)

SaldoInicial = 0
SaldoFinal = 0
Debito = 0
Credito = 0

'/////////////////////////////////BUSCO EL SALDO INICIAL///////////////////////////////////////////////////////////////////77
SQL = "SELECT MAX(Transacciones.CodCuentas) AS ÚltimoDeCodCuentas, MAX(Transacciones.NumeroMovimiento) AS ÚltimoDeNumeroMovimiento, MAX(Transacciones.NombreCuenta) AS ÚltimoDeNombreCuenta, Avg(Transacciones.TCambio) AS PromedioDeTCambio, Sum(Transacciones.TCambio*Transacciones.Debito) AS Debito, Sum(Transacciones.TCambio*Transacciones.Credito) AS Credito, Cuentas.KeyGrupo, Cuentas.DescripcionGrupo, Sum(Transacciones.TCambio*Transacciones.Debito)-Sum(Transacciones.TCambio*Transacciones.Credito) AS Saldo1, Sum(Transacciones.TCambio*Transacciones.Credito)-Sum(Transacciones.TCambio*Transacciones.Debito) AS Saldo2 " & _
"FROM (Grupos INNER JOIN Cuentas ON Grupos.KeyGrupo = Cuentas.KeyGrupo) INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas " & _
"Where (Transacciones.FechaTransaccion < CONVERT(DATETIME, '" & Fecha & "', 102)) GROUP BY Cuentas.KeyGrupo, Cuentas.DescripcionGrupo " & _
"Having (((Cuentas.KeyGrupo) = '" & llaveGrupo & "')) ORDER BY Cuentas.KeyGrupo "

'////////////////////////////////////BUSCO LOS MOVIMIENTOS DEL PERIODO//////////////////////////////////////////////////////


FrmReportes.DtaConsulta.RecordSource = SQL
FrmReportes.DtaConsulta.Refresh

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////BUSCO SI LAS CUENTAS TIENEN SALDO INICIAL/////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Not FrmReportes.DtaConsulta.Recordset.EOF Then
    If Tipo = "A" Or Tipo = "G" Or Tipo = "O" Then
               SaldoInicial = FrmReportes.DtaConsulta.Recordset("Saldo1")
               If SaldoInicial < 0 Then
                 Me.LblCreditoInicial.Caption = Format(Abs(SaldoInicial), "##,##0.00")
                 Me.LblDebitoInicial.Caption = "0.00"
                 TotalCreditoIni = TotalCreditoIni + SaldoInicial
               Else
                 Me.LblDebitoInicial.Caption = Format(SaldoInicial, "##,##0.00")
                 Me.LblCreditoInicial.Caption = "0.00"
                 TotalDebitoIni = TotalDebitoIni + SaldoInicial
               End If
               
'               Debito = Me.FldDebito.Text
'               Credito = Me.FldCredito.Text
          
               SaldoFinal = SaldoInicial + Debito - Credito
               If SaldoFinal < 0 Then
                          Me.LblCreditoFinal.Caption = Format(Abs(SaldoFinal), "##,##0.00")
                          Me.LblDebitoFinal.Caption = "0.00"
               Else
                          Me.LblDebitoFinal.Caption = Format(Abs(SaldoFinal), "##,##0.00")
                          Me.LblCreditoFinal.Caption = "0.00"
               End If
               Me.LblSaldoFinal.Caption = Format(SaldoFinal, "##,##0.00")
               Me.LblSaldoInicial.Caption = Format(SaldoInicial, "##,##0.00")
    Else
                 SaldoInicial = FrmReportes.DtaConsulta.Recordset("Saldo2")
                 Me.LblSaldoInicial.Caption = Format(SaldoInicial, "##,##0.00")
                 
                 If SaldoInicial < 0 Then
                   Me.LblDebitoInicial.Caption = Format(Abs(SaldoInicial), "##,##0.00")
                   TotalDebitoIni = TotalDebitoIni + SaldoInicial
                 Else
                   Me.LblCreditoInicial.Caption = Format(SaldoInicial, "##,##0.00")
                   TotalCreditoIni = TotalCreditoIni + SaldoInicial
                 End If
                   
            
'                 Debito = Me.FldDebito.Text
'                 Credito = Me.FldCredito.Text
                 
                 SaldoFinal = SaldoInicial + Credito - Debito
                 Me.LblSaldoFinal.Caption = Format(SaldoFinal, "##,##0.00")
                 
                 If SaldoFinal < 0 Then
                    Me.LblDebitoFinal.Caption = Format(Abs(SaldoFinal), "##,##0.00")
                 Else
                    Me.LblCreditoFinal.Caption = Format(Abs(SaldoFinal), "##,##0.00")
                 End If
     

    End If
Else

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////LAS CUENTAS NO TIENEN SALDO INICIAL///////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


     Me.LblDebitoInicial.Caption = "0.00"
     Me.LblCreditoInicial.Caption = "0.00"

     
     If Tipo = "A" Or Tipo = "G" Or Tipo = "O" Then
                 SaldoInicial = 0
                 
'                 Debito = Me.FldDebito.Text
'                 Credito = Me.FldCredito.Text
                 
                 SaldoFinal = SaldoInicial + Debito - Credito
                 
                 If SaldoFinal < 0 Then
                    Me.LblCreditoFinal.Caption = Format(Abs(SaldoFinal), "##,##0.00")
                 Else
                    Me.LblDebitoFinal.Caption = Format(SaldoFinal, "##,##0.00")
                 End If
            
                
                 Me.LblSaldoFinal.Caption = Format(SaldoFinal, "##,##0.00")
                 Me.LblSaldoInicial.Caption = Format(SaldoInicial, "##,##0.00")
    Else
                SaldoInicial = 0
'                If Me.FldDebito.Text = "" Then
'                 Debito = 0
'                Else
'                 Debito = Me.FldDebito.Text
'                End If
'
'                If Me.FldCredito.Text = "" Then
'                 Credito = 0
'                Else
'                Credito = Me.FldCredito.Text
'                End If
                SaldoFinal = SaldoInicial + Credito - Debito
                
                If SaldoFinal < 0 Then
                   Me.LblDebitoFinal.Caption = Format(Abs(SaldoFinal), "##,##0.00")
                Else
                   Me.LblCreditoFinal.Caption = Format(Abs(SaldoFinal), "##,##0.00")
                End If
                
                Me.LblSaldoFinal.Caption = Format(SaldoFinal, "##,##0.00")
                Me.LblSaldoInicial.Caption = Format(SaldoInicial, "##,##0.00")
      End If

End If

     '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
     '/////////////////////////////////////////////TOTALIZO EL REPORTE//////////////////////////////////////////////////////////
     '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
      TotalSaldoInicial = TotalSaldoInicial + SaldoInicial
      TotalSaldoFinal = TotalSaldoFinal + SaldoFinal
End Sub

Private Sub ActiveReport_ReportStart()
    On Error GoTo err
If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    End If
err:
    If err.Number <> 0 Then MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation
    
End Sub

Private Sub GroupFooter1_Format()
  Me.LblTotalSaldoInicial.Caption = Format(TotalSaldoInicial, "##,##0.00")
  Me.LblTotalSaldoFinal.Caption = Format(TotalSaldoFinal, "##,##0.00")
  
  Me.LblTotalCreditoInicial.Caption = Format(TotalCreditoIni, "##,##0.00")
  Me.LblTotalDebitoInicial.Caption = Format(TotalDebitoIni, "##,##0.00")
  
End Sub

