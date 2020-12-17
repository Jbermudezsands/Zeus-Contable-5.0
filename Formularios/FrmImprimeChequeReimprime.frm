VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmImprimeChequeReimprime 
   Caption         =   "Impresion de Cheques"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Encabezados"
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1935
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir Comprobante"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Imprimir Cheque y Comprobante"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Imprimir Cheque"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Consecutivo Cheque"
      Height          =   1815
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.TextBox LblConsecutivo 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmImprimeChequeReimprime.frx":0000
         TabIndex        =   2
         Top             =   840
         Width           =   855
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblCuenta 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "FrmImprimeChequeReimprime.frx":0070
      TabIndex        =   7
      Top             =   2040
      Width           =   4575
   End
End
Attribute VB_Name = "FrmImprimeChequeReimprime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ew As cls_NumEnglishWord
Private sw As cls_NumSpanishWord
Public Beneficiario As String, Memo As String, Monto As String, Letras As String, Consecutivo As String



Private Sub CmdImprimir_Click()
Dim Fechas1 As String, Fechas2 As String
Dim CodigoCuenta As String, Concepto As String
Dim x, y, H, V, Page As Integer, Dia As String, mes As String, Año As String
Dim i, J As Integer, Fechass As Date
Dim TotalDebito, TotalCredito, Totalpag As Double
Dim SubTotal, Total, IGV As Double, cadena As String
Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double, X5 As Double, Y5 As Double, X6 As Double, Y6 As Double, X7 As Double, Y7 As Double, X8 As Double, Y8 As Double, X9 As Double, Y9 As Double, X10 As Double, Y10 As Double, X11 As Double, Y11 As Double, X12 As Double, Y12 As Double, X13 As Double, Y13 As Double
Dim UltimaLinea As Double, DiferenciaY As Double, NLineas As Double
Dim Caracter As Double, ContadorLinea As Double, CadenaDescripcion As String, CaracteresLineas As Double
Dim Meses As Double, Monto As Double, TasaCambio As Double, FechaLetra As String, CaracteresConcepto As Double
Dim LineaConcepto As Double, MontoDolares As Double, MontoCordobas As Double, DebitoCordobas As Double, CreditoCordobas As Double, DebitoDolares As Double, CreditoDolares As Double
Dim TotalDebitoCordobas As Double, TotalCreditoCordobas As Double, TotalDebitoDolares As Double, TotalCreditoDolares As Double
Dim Ciudad As String



Page = 1

Printer.FontSize = 6
Printer.ScaleMode = 6



TotalDebito = 0
TotalCredito = 0

If Not Memo = "" Then
 Concepto = Memo
End If

If Not IsNumeric(Me.LblConsecutivo.Text) Then
 MsgBox "El Numero del Cheque debe Ser Numerico", vbCritical, "Sistema contable"
 Exit Sub
End If


'///////imprimo el reporte/////
 Debito = 0
 Credito = 0
 TotalDebito = 0
 TotalCredito = 0
      NumFecha1 = FrmCheque.TxtFecha.Value
      Fechas1 = Format(FrmCheque.TxtFecha.Value, "YYYY/MM/DD")
      NMovimiento = Val(FrmCheque.TxtNTransacciones)
      FrmCheque.DtaConsulta.RecordSource = "SELECT     FechaTransaccion, CodCuentas, NTransaccion, NumeroMovimiento, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, TCambio, Debito, Credito From Transacciones WHERE     (FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (NumeroMovimiento = " & NMovimiento & ")"
'      FrmCheque.DtaConsulta.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.CodCuentas, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, [Transacciones]![TCambio]*[Transacciones]![Debito] AS MDebito, [Transacciones]![TCambio]*[Transacciones]![Credito] AS MCredito, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito From Transacciones WHERE (((Transacciones.FechaTransaccion)=" & NumFecha1 & ") AND ((Transacciones.NumeroMovimiento)=" & NMovimiento & "))"
      FrmCheque.DtaConsulta.Refresh
      Do While Not FrmCheque.DtaConsulta.Recordset.EOF
      If FrmCheque.TxtMonto.Text = "" Then
       MontoCheque = 0
      Else
       MontoCheque = FrmCheque.TxtMonto
      End If
       Debito = FrmCheque.DtaConsulta.Recordset("Credito")
       Credito = FrmCheque.DtaConsulta.Recordset("Credito")
       TotalDebito = TotalDebito + Debito
       TotalCredito = TotalCredito + Credito
       FrmCheque.DtaConsulta.Recordset.MoveNext
      Loop
      
  CodigoCuenta = FrmCheque.DBCodigo.Text
  FrmCheque.DtaConsulta.RecordSource = "SELECT CodCuentas, ConsecutivoCheque From NConsecutivos WHERE (CodCuentas = '" & CodigoCuenta & "')"
  FrmCheque.DtaConsulta.Refresh
  If Not FrmCheque.DtaConsulta.Recordset.EOF Then
      FrmCheque.DtaConsulta.Recordset("ConsecutivoCheque") = Me.LblConsecutivo.Text
      FrmCheque.DtaConsulta.Recordset.Update
  End If
  
  
  FrmCheque.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento,Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito,Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas,Transacciones.NumeroMovimiento , Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas1 & "', 102))AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ")ORDER BY Transacciones.NTransaccion"
  FrmCheque.DtaConsulta.Refresh
  Do While Not FrmCheque.DtaConsulta.Recordset.EOF
  
    FrmCheque.DtaConsulta.Recordset("ChequeNo") = Me.LblConsecutivo.Text
    FrmCheque.DtaConsulta.Recordset.Update
    FrmCheque.DtaConsulta.Recordset.MoveNext
  Loop
  
If Me.Check1.Value = 1 Then

 ArepCheque2.DtaCheque.ConnectionString = ConexionReporte
 ArepCheque.DtaCheque.ConnectionString = ConexionReporte
 
 If FrmCheque.CmbMoneda.Text = "Córdobas" Then
   If FrmCheque.ChkCheque.Value = 1 Then
   
     TasaCambio = BuscaTasaCambio(FrmCheque.TxtFecha.Value)
     Monto = FrmCheque.TxtMonto.Text
     Monto = Monto / TasaCambio
     ArepCheque.LblDescripcionMonto.Caption = sw.ConvertCurrencyToSpanish(Monto, "Dólares")
     ArepCheque.LblMonto.Caption = Format(Monto, "##,##0.00")
     
     ArepCheque2.LblDescripcionMonto.Caption = sw.ConvertCurrencyToSpanish(Monto, "Dólares")
     ArepCheque2.LblMonto.Caption = Format(Monto, "##,##0.00")
   Else
    ArepCheque.LblDescripcionMonto.Caption = FrmCheque.TxtLetras.Text
    ArepCheque.LblMonto.Caption = Format(FrmCheque.TxtMonto.Text, "##,##0.00")
    
    ArepCheque2.LblDescripcionMonto.Caption = FrmCheque.TxtLetras.Text
    ArepCheque2.LblMonto.Caption = Format(FrmCheque.TxtMonto.Text, "##,##0.00")
   End If
 Else
    ArepCheque.LblDescripcionMonto.Caption = FrmCheque.TxtLetras.Text
    ArepCheque.LblMonto.Caption = Format(FrmCheque.TxtMonto.Text, "##,##0.00")
    
    ArepCheque2.LblDescripcionMonto.Caption = FrmCheque.TxtLetras.Text
    ArepCheque2.LblMonto.Caption = Format(FrmCheque.TxtMonto.Text, "##,##0.00")
 End If
 ArepCheque.LblMemo.Caption = FrmCheque.TxtMemo.Text
 ArepCheque2.LblMemo.Caption = FrmCheque.TxtMemo.Text
 
 ArepCheque.LblNombre.Caption = FrmCheque.TxtNombre.Text
 ArepCheque.LblChequeNo.Caption = Me.LblConsecutivo.Text
 
 ArepCheque2.LblNombre.Caption = FrmCheque.TxtNombre.Text
 ArepCheque2.LblChequeNo.Caption = Me.LblConsecutivo.Text

FechaCheque = Fechas1
NumeroMovimientos = NumeroTransaccion

    If FrmCheque.CmbMoneda.Text = "Córdobas" Then
        
        ArepCheque.DtaCheque.Source = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NumeroMovimiento, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.Debito, Transacciones.Credito, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTasas,  " & _
                                      "CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito / Tasas.MontoCordobas ELSE Transacciones.Debito END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito / Tasas.MontoCordobas ELSE Transacciones.Credito END AS CreditoD, Transacciones.NPeriodo, Cuentas.DescripcionGrupo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
                                      "WHERE (Transacciones.FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ") ORDER BY Transacciones.NTransaccion"
        ArepCheque2.DtaCheque.Source = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NumeroMovimiento, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.Debito, Transacciones.Credito, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTasas,  " & _
                                      "CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito / Tasas.MontoCordobas ELSE Transacciones.Debito END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito / Tasas.MontoCordobas ELSE Transacciones.Credito END AS CreditoD, Transacciones.NPeriodo, Cuentas.DescripcionGrupo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
                                      "WHERE (Transacciones.FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ") ORDER BY Transacciones.NTransaccion"
    Else
        ArepCheque.DtaCheque.Source = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NumeroMovimiento, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.Debito, Transacciones.Credito, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTasas,  " & _
                                      "CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito * Tasas.MontoCordobas ELSE Transacciones.Debito END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito / Tasas.MontoCordobas ELSE Transacciones.Credito END AS CreditoD, Transacciones.NPeriodo, Cuentas.DescripcionGrupo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
                                      "WHERE (Transacciones.FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ") ORDER BY Transacciones.NTransaccion"
        ArepCheque2.DtaCheque.Source = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NumeroMovimiento, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito*Tasas.MontoCordobas  ELSE Transacciones.Debito END AS Debito,  CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito*Tasas.MontoCordobas  ELSE Transacciones.Credito END AS Credito, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTasas,  " & _
                                      "CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito  ELSE Transacciones.Debito * Tasas.MontoCordobas END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito  ELSE Transacciones.Credito * Tasas.MontoCordobas END AS CreditoD, Transacciones.NPeriodo, Cuentas.DescripcionGrupo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
                                      "WHERE (Transacciones.FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ") ORDER BY Transacciones.NTransaccion"
    
    End If
 ArepCheque2.Show 1
 
ElseIf Me.Check2.Value = 1 Then

'---------------------------------------------------------------------------------------------------------------------------------------
'-------------------------------------IMPRIMO EL COMPROBANTE PRIMERO------------------------------------
'-------------------------------------------------------------------------------------------------------
         ArepCheque2.DtaCheque.ConnectionString = ConexionReporte
         ArepCheque.DtaCheque.ConnectionString = ConexionReporte
         
         If FrmCheque.CmbMoneda.Text = "Córdobas" Then
           If FrmCheque.ChkCheque.Value = 1 Then
           
             TasaCambio = BuscaTasaCambio(FrmCheque.TxtFecha.Value)
             Monto = FrmCheque.TxtMonto.Text
             Monto = Monto / TasaCambio
             ArepCheque.LblDescripcionMonto.Caption = sw.ConvertCurrencyToSpanish(Monto, "Dólares")
             ArepCheque.LblMonto.Caption = Format(Monto, "##,##0.00")
             
             ArepCheque2.LblDescripcionMonto.Caption = sw.ConvertCurrencyToSpanish(Monto, "Dólares")
             ArepCheque2.LblMonto.Caption = Format(Monto, "##,##0.00")
           Else
            ArepCheque.LblDescripcionMonto.Caption = FrmCheque.TxtLetras.Text
            ArepCheque.LblMonto.Caption = Format(FrmCheque.TxtMonto.Text, "##,##0.00")
            
            ArepCheque2.LblDescripcionMonto.Caption = FrmCheque.TxtLetras.Text
            ArepCheque2.LblMonto.Caption = Format(FrmCheque.TxtMonto.Text, "##,##0.00")
           End If
         Else
            ArepCheque.LblDescripcionMonto.Caption = FrmCheque.TxtLetras.Text
            ArepCheque.LblMonto.Caption = Format(FrmCheque.TxtMonto.Text, "##,##0.00")
            
            ArepCheque2.LblDescripcionMonto.Caption = FrmCheque.TxtLetras.Text
            ArepCheque2.LblMonto.Caption = Format(FrmCheque.TxtMonto.Text, "##,##0.00")
         End If
         ArepCheque.LblMemo.Caption = FrmCheque.TxtMemo.Text
         ArepCheque2.LblMemo.Caption = FrmCheque.TxtMemo.Text
         
         ArepCheque.LblNombre.Caption = FrmCheque.TxtNombre.Text
         ArepCheque.LblChequeNo.Caption = Me.LblConsecutivo.Text
         
         ArepCheque2.LblNombre.Caption = FrmCheque.TxtNombre.Text
         ArepCheque2.LblChequeNo.Caption = Me.LblConsecutivo.Text
        
        FechaCheque = Fechas1
        NumeroMovimientos = NumeroTransaccion
        
            If FrmCheque.CmbMoneda.Text = "Córdobas" Then
                
                ArepCheque.DtaCheque.Source = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NumeroMovimiento, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.Debito, Transacciones.Credito, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTasas,  " & _
                                              "CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito / Tasas.MontoCordobas ELSE Transacciones.Debito END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito / Tasas.MontoCordobas ELSE Transacciones.Credito END AS CreditoD, Transacciones.NPeriodo, Cuentas.DescripcionGrupo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
                                              "WHERE (Transacciones.FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ") ORDER BY Transacciones.NTransaccion"
                ArepCheque2.DtaCheque.Source = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NumeroMovimiento, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.Debito, Transacciones.Credito, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTasas,  " & _
                                              "CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito / Tasas.MontoCordobas ELSE Transacciones.Debito END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito / Tasas.MontoCordobas ELSE Transacciones.Credito END AS CreditoD, Transacciones.NPeriodo, Cuentas.DescripcionGrupo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
                                              "WHERE (Transacciones.FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ") ORDER BY Transacciones.NTransaccion"
            Else
                ArepCheque.DtaCheque.Source = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NumeroMovimiento, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.Debito, Transacciones.Credito, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTasas,  " & _
                                              "CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito * Tasas.MontoCordobas ELSE Transacciones.Debito END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito / Tasas.MontoCordobas ELSE Transacciones.Credito END AS CreditoD, Transacciones.NPeriodo, Cuentas.DescripcionGrupo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
                                              "WHERE (Transacciones.FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ") ORDER BY Transacciones.NTransaccion"
                ArepCheque2.DtaCheque.Source = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NumeroMovimiento, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito*Tasas.MontoCordobas  ELSE Transacciones.Debito END AS Debito,  CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito*Tasas.MontoCordobas  ELSE Transacciones.Credito END AS Credito, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTasas,  " & _
                                              "CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito  ELSE Transacciones.Debito * Tasas.MontoCordobas END AS DebitoD, CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito  ELSE Transacciones.Credito * Tasas.MontoCordobas END AS CreditoD, Transacciones.NPeriodo, Cuentas.DescripcionGrupo FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento " & _
                                              "WHERE (Transacciones.FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ") ORDER BY Transacciones.NTransaccion"
            
            End If
         ArepCheque2.Show 1
 
 
        MsgBox "Coloque El Comprobante en la impresora", vbInformation, "Zeus Contable"
        
 
        FrmCheque.AdoCordenadas.RecordSource = "SELECT CodCuenta, X1, Y1, X2, Y2, X3, Y3, X4, Y4, X5, Y5, X6, Y6, X7, Y7, X8, Y8, X9, Y9, X10, Y10, X11, Y11, X12, Y12, X13, Y13,X14, Y14,X15, Y15,X16, Y16,X17, Y17, X18, Y18, X19, Y19,X20, Y20,X21, Y21, X22, Y22, NLineas,CaracteresLineas, CaracteresConcepto, Ciudad From CordenadasCheque WHERE  (CodCuenta = '" & CodigoCuenta & "')"
        FrmCheque.AdoCordenadas.Refresh
        If FrmCheque.AdoCordenadas.Recordset.EOF Then
         MsgBox "No Existen Coordenadas, para la Cuenta", vbCritical, "Sistema Contable"
         Exit Sub
        End If


        X1 = FrmCheque.AdoCordenadas.Recordset("X1")
        Y1 = FrmCheque.AdoCordenadas.Recordset("Y1")
        X2 = FrmCheque.AdoCordenadas.Recordset("X2")
        Y2 = FrmCheque.AdoCordenadas.Recordset("Y2")
        X3 = FrmCheque.AdoCordenadas.Recordset("X3")
        Y3 = FrmCheque.AdoCordenadas.Recordset("Y3")
        X4 = FrmCheque.AdoCordenadas.Recordset("X4")
        Y4 = FrmCheque.AdoCordenadas.Recordset("Y4")
        X5 = FrmCheque.AdoCordenadas.Recordset("X5")
        Y5 = FrmCheque.AdoCordenadas.Recordset("Y5")
        X6 = FrmCheque.AdoCordenadas.Recordset("X6")
        Y6 = FrmCheque.AdoCordenadas.Recordset("Y6")
        X7 = FrmCheque.AdoCordenadas.Recordset("X7")
        Y7 = FrmCheque.AdoCordenadas.Recordset("Y7")
        X8 = FrmCheque.AdoCordenadas.Recordset("X8")
        Y8 = FrmCheque.AdoCordenadas.Recordset("Y8")
        X9 = FrmCheque.AdoCordenadas.Recordset("X9")
        Y9 = FrmCheque.AdoCordenadas.Recordset("Y9")
        X10 = FrmCheque.AdoCordenadas.Recordset("X10")
        Y10 = FrmCheque.AdoCordenadas.Recordset("Y10")
        X11 = FrmCheque.AdoCordenadas.Recordset("X11")
        Y11 = FrmCheque.AdoCordenadas.Recordset("Y11")
        X12 = FrmCheque.AdoCordenadas.Recordset("X12")
        Y12 = FrmCheque.AdoCordenadas.Recordset("Y12")
        X13 = FrmCheque.AdoCordenadas.Recordset("X13")
        Y13 = FrmCheque.AdoCordenadas.Recordset("Y13")
        X14 = FrmCheque.AdoCordenadas.Recordset("X14")
        Y14 = FrmCheque.AdoCordenadas.Recordset("Y14")
        X15 = FrmCheque.AdoCordenadas.Recordset("X15")
        Y15 = FrmCheque.AdoCordenadas.Recordset("Y15")
        X16 = FrmCheque.AdoCordenadas.Recordset("X16")
        Y16 = FrmCheque.AdoCordenadas.Recordset("Y16")
        X17 = FrmCheque.AdoCordenadas.Recordset("X17")
        Y17 = FrmCheque.AdoCordenadas.Recordset("Y17")
        X18 = FrmCheque.AdoCordenadas.Recordset("X18")
        Y18 = FrmCheque.AdoCordenadas.Recordset("Y18")
        X19 = FrmCheque.AdoCordenadas.Recordset("X19")
        Y19 = FrmCheque.AdoCordenadas.Recordset("Y19")
        X20 = FrmCheque.AdoCordenadas.Recordset("X20")
        Y20 = FrmCheque.AdoCordenadas.Recordset("Y20")
        X21 = FrmCheque.AdoCordenadas.Recordset("X21")
        Y21 = FrmCheque.AdoCordenadas.Recordset("Y21")
        X22 = FrmCheque.AdoCordenadas.Recordset("X22")
        Y22 = FrmCheque.AdoCordenadas.Recordset("Y22")
        NLineas = Val(FrmCheque.AdoCordenadas.Recordset("NLineas"))
        CaracteresLineas = Val(FrmCheque.AdoCordenadas.Recordset("CaracteresLineas"))
        CaracteresConcepto = Val(FrmCheque.AdoCordenadas.Recordset("CaracteresConcepto"))
        Ciudad = FrmCheque.AdoCordenadas.Recordset("Ciudad")
        
       If FrmCheque.CmbMoneda.Text = "Córdobas" Then
            If FrmCheque.ChkCheque.Value = 1 Then
              TasaCambio = BuscaTasaCambio(FrmCheque.TxtFecha.Value)
              Monto = FrmCheque.TxtMonto.Text
              Monto = Monto / TasaCambio
              FrmCheque.TxtLetras.Text = sw.ConvertCurrencyToSpanish(Monto, "Dólares")
'              ArepCheque.LblMonto.Caption = Format(Monto, "##,##0.00")
            Else
              Monto = FrmCheque.TxtMonto.Text
            End If
         Else
           Monto = FrmCheque.TxtMonto.Text
      End If
      
      
        FrmCheque.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento,Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito,Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas,Transacciones.NumeroMovimiento , Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas1 & "', 102))AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ")ORDER BY Transacciones.NTransaccion"
        FrmCheque.DtaConsulta.Refresh
        Printer.FontSize = 8
        'Inicio el Ciclo de Impresion
        i = 1
        
        TasaCambio = BuscaTasaCambio(FrmCheque.TxtFecha.Value)
'        Do While Not FrmCheque.DtaConsulta.Recordset.EOF
        
                   If i = 1 Then
                        ContadorLinea = i
                          
                                  '//////////////////////////////////////////////////////////////////////////////////////
                                  '//////////////////////////IMPRIMO LOS ENCABEZADOS ////////////////////////////////////
                                  '//////////////////////////////////////////////////////////////////////////////////
                           
                                   If X5 <> 0 Or Y5 <> 0 Then
                                     Caracter = 1
                                     LineaConcepto = 1
                                     cadena = Concepto
                                     If Len(cadena) > CaracteresConcepto Then
                                          Do While Len(cadena) >= CaracteresConcepto
                                                 If Caracter = 1 Then
                '                                    Printer.CurrentX = Val(X5) '5
                '                                    Printer.CurrentY = Val(Y5) + (5 * i) '120
                '                                    Printer.FontName = "Times New Roman"
                '                                    Printer.FontSize = 11
                '                                    Printer.FontBold = True
                '                                    Printer.Print Concepto
                                                    
                                                           
                                                                 cadena = Mid(Concepto, 1, CaracteresConcepto)
                                                                 Printer.CurrentX = Val(X5) '25
                                                                 Printer.CurrentY = Val(Y5) + (5 * LineaConcepto)
                                                                 Printer.FontName = "Times New Roman"
                                                                 Printer.FontSize = 11
                                                                 Printer.FontBold = True
                                                                 Printer.Print cadena
                                                                 Caracter = Caracter + CaracteresConcepto
                                                                 
                                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA SIGUIENTE/////////////////
                                                                 
                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                 If Len(cadena) < CaracteresConcepto Then
                                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                                     LineaConcepto = LineaConcepto + 1
                                                                     Printer.CurrentX = Val(X5) '25
                                                                     Printer.CurrentY = Val(Y5) + (5 * LineaConcepto)
                                                                     Printer.FontName = "Times New Roman"
                                                                     Printer.FontSize = 11
                                                                     Printer.FontBold = True
                                                                     Printer.Print cadena
                                                                     
                                                                     Caracter = Caracter + CaracteresConcepto
                                                                 End If
                                                                 
                                                 Else
                                                                 
                                                                 LineaConcepto = LineaConcepto + 1
                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                 Printer.CurrentX = Val(X5) '25
                                                                 Printer.CurrentY = Val(Y5) + (5 * LineaConcepto)
                                                                 Printer.FontName = "Times New Roman"
                                                                 Printer.FontSize = 11
                                                                 Printer.FontBold = True
                                                                 Printer.Print cadena
                                                                 
                                                                 Caracter = Caracter + CaracteresConcepto
                                                                 
                                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA/////////////////
                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                 If Len(cadena) < CaracteresConcepto Then
                                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                                     LineaConcepto = LineaConcepto + 1
                                                                     Printer.CurrentX = Val(X5) '25
                                                                     Printer.CurrentY = Val(Y5) + (5 * LineaConcepto)
                                                                     Printer.FontName = "Times New Roman"
                                                                     Printer.FontSize = 11
                                                                     Printer.FontBold = True
                                                                     Printer.Print cadena
                                                                     
                                                                     Caracter = Caracter + CaracteresConcepto
                                                                 End If
                                                    
                                                 End If
                                          Loop
                                          
                                     Else
                                                    Printer.CurrentX = Val(X5) '5
                                                    Printer.CurrentY = Val(Y5) + (5 * i) '120
                                                    Printer.FontName = "Times New Roman"
                                                    Printer.FontSize = 11
                                                    Printer.FontBold = True
                                                    Printer.Print Concepto
                                     End If
                                   End If
                                   
                                   
                                  If X18 <> 0 Or Y18 <> 0 Then
                                     Caracter = 1
                                     LineaConcepto = 1
                                     cadena = Concepto
                                     If Len(cadena) > CaracteresConcepto Then
                                          Do While Len(cadena) >= CaracteresConcepto
                                                 If Caracter = 1 Then
                '                                    Printer.CurrentX = Val(X5) '5
                '                                    Printer.CurrentY = Val(Y5) + (5 * i) '120
                '                                    Printer.FontName = "Times New Roman"
                '                                    Printer.FontSize = 11
                '                                    Printer.FontBold = True
                '                                    Printer.Print Concepto
                                                    
                                                           
                                                                 cadena = Mid(Concepto, 1, CaracteresConcepto)
                                                                 Printer.CurrentX = Val(X18) '25
                                                                 Printer.CurrentY = Val(Y18) + (5 * LineaConcepto)
                                                                 Printer.FontName = "Times New Roman"
                                                                 Printer.FontSize = 11
                                                                 Printer.FontBold = True
                                                                 Printer.Print cadena
                                                                 Caracter = Caracter + CaracteresConcepto
                                                                 
                                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA SIGUIENTE/////////////////
                                                                 
                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                 If Len(cadena) < CaracteresConcepto Then
                                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                                     LineaConcepto = LineaConcepto + 1
                                                                     Printer.CurrentX = Val(X18) '25
                                                                     Printer.CurrentY = Val(Y18) + (5 * LineaConcepto)
                                                                     Printer.FontName = "Times New Roman"
                                                                     Printer.FontSize = 11
                                                                     Printer.FontBold = True
                                                                     Printer.Print cadena
                                                                     
                                                                     Caracter = Caracter + CaracteresConcepto
                                                                 End If
                                                                 
                                                 Else
                                                                 
                                                                 LineaConcepto = LineaConcepto + 1
                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                 Printer.CurrentX = Val(X18) '25
                                                                 Printer.CurrentY = Val(Y18) + (5 * LineaConcepto)
                                                                 Printer.FontName = "Times New Roman"
                                                                 Printer.FontSize = 11
                                                                 Printer.FontBold = True
                                                                 Printer.Print cadena
                                                                 
                                                                 Caracter = Caracter + CaracteresConcepto
                                                                 
                                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA/////////////////
                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                 If Len(cadena) < CaracteresConcepto Then
                                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                                     LineaConcepto = LineaConcepto + 1
                                                                     Printer.CurrentX = Val(X18) '25
                                                                     Printer.CurrentY = Val(Y18) + (5 * LineaConcepto)
                                                                     Printer.FontName = "Times New Roman"
                                                                     Printer.FontSize = 11
                                                                     Printer.FontBold = True
                                                                     Printer.Print cadena
                                                                     
                                                                     Caracter = Caracter + CaracteresConcepto
                                                                 End If
                                                    
                                                 End If
                                          Loop
                                          
                                     Else
                                                    Printer.CurrentX = Val(X18) '5
                                                    Printer.CurrentY = Val(Y18) + (5 * i) '120
                                                    Printer.FontName = "Times New Roman"
                                                    Printer.FontSize = 11
                                                    Printer.FontBold = True
                                                    Printer.Print Concepto
                                     End If
                                   End If
                                    
                                    Dia = Day(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                    mes = Month(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                    Año = Year(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                    Meses = Month(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                  
                                '    FrmCheque.DtaConsulta.Recordset.MoveLast
                                   If X9 <> 0 Or Y9 <> 0 Then
                                    Printer.CurrentX = X9
                                    Printer.CurrentY = Y9
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 11
                                    Printer.FontBold = True
                '                    Printer.Print FrmCheque.DtaConsulta.Recordset("NumeroMovimiento")
                                    Printer.Print Me.LblConsecutivo.Text
                                   End If
                                    
                                   If X1 <> 0 Or Y1 <> 0 Then
                                    Printer.CurrentX = Val(X1)
                                    Printer.CurrentY = Val(Y1) + (5 * i)
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 9
                                    Printer.FontBold = True
                                    Printer.Print FrmCheque.TxtNombre.Text
                                   End If
                                   
                                  If X14 <> 0 Or Y14 <> 0 Then
                                    Printer.CurrentX = Val(X14)
                                    Printer.CurrentY = Val(Y14) + (5 * i)
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 9
                                    Printer.FontBold = True
                                    Printer.Print FrmCheque.TxtNombre.Text
                                   End If
                                   
                                   If X4 <> 0 Or Y4 <> 0 Then
                                    Printer.CurrentX = Val(X4)
                                    Printer.CurrentY = Val(Y4) + (5 * i)
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 9
                                    Printer.FontBold = True
                                    Printer.Print FrmCheque.TxtLetras.Text
                                   End If
                                   
                                  If X15 <> 0 Or Y15 <> 0 Then
                                    Printer.CurrentX = Val(X15)
                                    Printer.CurrentY = Val(Y15) + (5 * i)
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 9
                                    Printer.FontBold = True
                                    Printer.Print FrmCheque.TxtLetras.Text
                                   End If
                                   
                                   If X3 <> 0 Or Y3 <> 0 Then
                                    Printer.CurrentX = Val(X3)
                                    Printer.CurrentY = Val(Y3) + (5 * i)
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 11
                                    Printer.FontBold = True
                                    Printer.Print Format(Monto, "##,##0.00")
                                   End If
                                   
                                                      
                                   If X16 <> 0 Or Y16 <> 0 Then
                                    Printer.CurrentX = Val(X16)
                                    Printer.CurrentY = Val(Y16) + (5 * i)
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 11
                                    Printer.FontBold = True
                                    Printer.Print Format(Monto, "##,##0.00")
                                   End If
                                
                                   If X2 <> 0 Or Y2 <> 0 Then
                                    Printer.CurrentX = Val(X2) '20
                                    Printer.CurrentY = Val(Y2) '288
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 11
                                    Printer.FontBold = True
                                    FechaLetra = Ciudad & "          " & Format(Day(FrmCheque.DtaConsulta.Recordset("FechaTransaccion")), "0#") & "          " & Format(Month(FrmCheque.DtaConsulta.Recordset("FechaTransaccion")), "0#") & "           " & Year(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                    Printer.Print FechaLetra
                                   End If
                                
                                    If X17 <> 0 Or Y17 <> 0 Then
                                    Printer.CurrentX = Val(X17) '20
                                    Printer.CurrentY = Val(Y17) '288
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 11
                                    Printer.FontBold = True
                                    FechaLetra = Ciudad & "          " & Format(Day(FrmCheque.DtaConsulta.Recordset("FechaTransaccion")), "0#") & "          " & Format(Month(FrmCheque.DtaConsulta.Recordset("FechaTransaccion")), "0#") & "           " & Year(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                    Printer.Print FechaLetra
                                   End If
                                
                   End If

        
        Printer.EndDoc
        
'         FrmCheque.DtaConsulta.Recordset.MoveNext
'        Loop

ElseIf Me.Check3.Value = 1 Then

        FrmCheque.AdoCordenadas.RecordSource = "SELECT CodCuenta, X1, Y1, X2, Y2, X3, Y3, X4, Y4, X5, Y5, X6, Y6, X7, Y7, X8, Y8, X9, Y9, X10, Y10, X11, Y11, X12, Y12, X13, Y13,X14, Y14,X15, Y15,X16, Y16,X17, Y17, X18, Y18, X19, Y19,X20, Y20,X21, Y21, X22, Y22, NLineas,CaracteresLineas, CaracteresConcepto From CordenadasCheque WHERE  (CodCuenta = '" & CodigoCuenta & "')"
        FrmCheque.AdoCordenadas.Refresh
        If FrmCheque.AdoCordenadas.Recordset.EOF Then
         MsgBox "No Existen Coordenadas, para la Cuenta", vbCritical, "Sistema Contable"
         Exit Sub
        End If


        X1 = FrmCheque.AdoCordenadas.Recordset("X1")
        Y1 = FrmCheque.AdoCordenadas.Recordset("Y1")
        X2 = FrmCheque.AdoCordenadas.Recordset("X2")
        Y2 = FrmCheque.AdoCordenadas.Recordset("Y2")
        X3 = FrmCheque.AdoCordenadas.Recordset("X3")
        Y3 = FrmCheque.AdoCordenadas.Recordset("Y3")
        X4 = FrmCheque.AdoCordenadas.Recordset("X4")
        Y4 = FrmCheque.AdoCordenadas.Recordset("Y4")
        X5 = FrmCheque.AdoCordenadas.Recordset("X5")
        Y5 = FrmCheque.AdoCordenadas.Recordset("Y5")
        X6 = FrmCheque.AdoCordenadas.Recordset("X6")
        Y6 = FrmCheque.AdoCordenadas.Recordset("Y6")
        X7 = FrmCheque.AdoCordenadas.Recordset("X7")
        Y7 = FrmCheque.AdoCordenadas.Recordset("Y7")
        X8 = FrmCheque.AdoCordenadas.Recordset("X8")
        Y8 = FrmCheque.AdoCordenadas.Recordset("Y8")
        X9 = FrmCheque.AdoCordenadas.Recordset("X9")
        Y9 = FrmCheque.AdoCordenadas.Recordset("Y9")
        X10 = FrmCheque.AdoCordenadas.Recordset("X10")
        Y10 = FrmCheque.AdoCordenadas.Recordset("Y10")
        X11 = FrmCheque.AdoCordenadas.Recordset("X11")
        Y11 = FrmCheque.AdoCordenadas.Recordset("Y11")
        X12 = FrmCheque.AdoCordenadas.Recordset("X12")
        Y12 = FrmCheque.AdoCordenadas.Recordset("Y12")
        X13 = FrmCheque.AdoCordenadas.Recordset("X13")
        Y13 = FrmCheque.AdoCordenadas.Recordset("Y13")
        X14 = FrmCheque.AdoCordenadas.Recordset("X14")
        Y14 = FrmCheque.AdoCordenadas.Recordset("Y14")
        X15 = FrmCheque.AdoCordenadas.Recordset("X15")
        Y15 = FrmCheque.AdoCordenadas.Recordset("Y15")
        X16 = FrmCheque.AdoCordenadas.Recordset("X16")
        Y16 = FrmCheque.AdoCordenadas.Recordset("Y16")
        X17 = FrmCheque.AdoCordenadas.Recordset("X17")
        Y17 = FrmCheque.AdoCordenadas.Recordset("Y17")
        X18 = FrmCheque.AdoCordenadas.Recordset("X18")
        Y18 = FrmCheque.AdoCordenadas.Recordset("Y18")
        X19 = FrmCheque.AdoCordenadas.Recordset("X19")
        Y19 = FrmCheque.AdoCordenadas.Recordset("Y19")
        X20 = FrmCheque.AdoCordenadas.Recordset("X20")
        Y20 = FrmCheque.AdoCordenadas.Recordset("Y20")
        X21 = FrmCheque.AdoCordenadas.Recordset("X21")
        Y21 = FrmCheque.AdoCordenadas.Recordset("Y21")
        X22 = FrmCheque.AdoCordenadas.Recordset("X22")
        Y22 = FrmCheque.AdoCordenadas.Recordset("Y22")
        NLineas = Val(FrmCheque.AdoCordenadas.Recordset("NLineas"))
        CaracteresLineas = Val(FrmCheque.AdoCordenadas.Recordset("CaracteresLineas"))
        CaracteresConcepto = Val(FrmCheque.AdoCordenadas.Recordset("CaracteresConcepto"))
        
       If FrmCheque.CmbMoneda.Text = "Córdobas" Then
            If FrmCheque.ChkCheque.Value = 1 Then
              TasaCambio = BuscaTasaCambio(FrmCheque.TxtFecha.Value)
              Monto = FrmCheque.TxtMonto.Text
              Monto = Monto / TasaCambio
              FrmCheque.TxtLetras.Text = sw.ConvertCurrencyToSpanish(Monto, "Dólares")
'              ArepCheque.LblMonto.Caption = Format(Monto, "##,##0.00")
            Else
              Monto = FrmCheque.TxtMonto.Text
            End If
         Else
           Monto = FrmCheque.TxtMonto.Text
      End If
      
      
        FrmCheque.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento,Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito,Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas,Transacciones.NumeroMovimiento , Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas1 & "', 102))AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ")ORDER BY Transacciones.NTransaccion"
        FrmCheque.DtaConsulta.Refresh
        Printer.FontSize = 8
        'Inicio el Ciclo de Impresion
        i = 1
        
        TasaCambio = BuscaTasaCambio(FrmCheque.TxtFecha.Value)
'        Do While Not FrmCheque.DtaConsulta.Recordset.EOF
        
                   If i = 1 Then
                        ContadorLinea = i
                          
                                  '//////////////////////////////////////////////////////////////////////////////////////
                                  '//////////////////////////IMPRIMO LOS ENCABEZADOS ////////////////////////////////////
                                  '//////////////////////////////////////////////////////////////////////////////////
                           
                                   If X5 <> 0 Or Y5 <> 0 Then
                                     Caracter = 1
                                     LineaConcepto = 1
                                     cadena = Concepto
                                     If Len(cadena) > CaracteresConcepto Then
                                          Do While Len(cadena) >= CaracteresConcepto
                                                 If Caracter = 1 Then
                '                                    Printer.CurrentX = Val(X5) '5
                '                                    Printer.CurrentY = Val(Y5) + (5 * i) '120
                '                                    Printer.FontName = "Times New Roman"
                '                                    Printer.FontSize = 11
                '                                    Printer.FontBold = True
                '                                    Printer.Print Concepto
                                                    
                                                           
                                                                 cadena = Mid(Concepto, 1, CaracteresConcepto)
                                                                 Printer.CurrentX = Val(X5) '25
                                                                 Printer.CurrentY = Val(Y5) + (5 * LineaConcepto)
                                                                 Printer.FontName = "Times New Roman"
                                                                 Printer.FontSize = 11
                                                                 Printer.FontBold = True
                                                                 Printer.Print cadena
                                                                 Caracter = Caracter + CaracteresConcepto
                                                                 
                                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA SIGUIENTE/////////////////
                                                                 
                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                 If Len(cadena) < CaracteresConcepto Then
                                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                                     LineaConcepto = LineaConcepto + 1
                                                                     Printer.CurrentX = Val(X5) '25
                                                                     Printer.CurrentY = Val(Y5) + (5 * LineaConcepto)
                                                                     Printer.FontName = "Times New Roman"
                                                                     Printer.FontSize = 11
                                                                     Printer.FontBold = True
                                                                     Printer.Print cadena
                                                                     
                                                                     Caracter = Caracter + CaracteresConcepto
                                                                 End If
                                                                 
                                                 Else
                                                                 
                                                                 LineaConcepto = LineaConcepto + 1
                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                 Printer.CurrentX = Val(X5) '25
                                                                 Printer.CurrentY = Val(Y5) + (5 * LineaConcepto)
                                                                 Printer.FontName = "Times New Roman"
                                                                 Printer.FontSize = 11
                                                                 Printer.FontBold = True
                                                                 Printer.Print cadena
                                                                 
                                                                 Caracter = Caracter + CaracteresConcepto
                                                                 
                                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA/////////////////
                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                 If Len(cadena) < CaracteresConcepto Then
                                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                                     LineaConcepto = LineaConcepto + 1
                                                                     Printer.CurrentX = Val(X5) '25
                                                                     Printer.CurrentY = Val(Y5) + (5 * LineaConcepto)
                                                                     Printer.FontName = "Times New Roman"
                                                                     Printer.FontSize = 11
                                                                     Printer.FontBold = True
                                                                     Printer.Print cadena
                                                                     
                                                                     Caracter = Caracter + CaracteresConcepto
                                                                 End If
                                                    
                                                 End If
                                          Loop
                                          
                                     Else
                                                    Printer.CurrentX = Val(X5) '5
                                                    Printer.CurrentY = Val(Y5) + (5 * i) '120
                                                    Printer.FontName = "Times New Roman"
                                                    Printer.FontSize = 11
                                                    Printer.FontBold = True
                                                    Printer.Print Concepto
                                     End If
                                   End If
                                   
                                   
                                  If X18 <> 0 Or Y18 <> 0 Then
                                     Caracter = 1
                                     LineaConcepto = 1
                                     cadena = Concepto
                                     If Len(cadena) > CaracteresConcepto Then
                                          Do While Len(cadena) >= CaracteresConcepto
                                                 If Caracter = 1 Then
                '                                    Printer.CurrentX = Val(X5) '5
                '                                    Printer.CurrentY = Val(Y5) + (5 * i) '120
                '                                    Printer.FontName = "Times New Roman"
                '                                    Printer.FontSize = 11
                '                                    Printer.FontBold = True
                '                                    Printer.Print Concepto
                                                    
                                                           
                                                                 cadena = Mid(Concepto, 1, CaracteresConcepto)
                                                                 Printer.CurrentX = Val(X18) '25
                                                                 Printer.CurrentY = Val(Y18) + (5 * LineaConcepto)
                                                                 Printer.FontName = "Times New Roman"
                                                                 Printer.FontSize = 11
                                                                 Printer.FontBold = True
                                                                 Printer.Print cadena
                                                                 Caracter = Caracter + CaracteresConcepto
                                                                 
                                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA SIGUIENTE/////////////////
                                                                 
                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                 If Len(cadena) < CaracteresConcepto Then
                                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                                     LineaConcepto = LineaConcepto + 1
                                                                     Printer.CurrentX = Val(X18) '25
                                                                     Printer.CurrentY = Val(Y18) + (5 * LineaConcepto)
                                                                     Printer.FontName = "Times New Roman"
                                                                     Printer.FontSize = 11
                                                                     Printer.FontBold = True
                                                                     Printer.Print cadena
                                                                     
                                                                     Caracter = Caracter + CaracteresConcepto
                                                                 End If
                                                                 
                                                 Else
                                                                 
                                                                 LineaConcepto = LineaConcepto + 1
                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                 Printer.CurrentX = Val(X18) '25
                                                                 Printer.CurrentY = Val(Y18) + (5 * LineaConcepto)
                                                                 Printer.FontName = "Times New Roman"
                                                                 Printer.FontSize = 11
                                                                 Printer.FontBold = True
                                                                 Printer.Print cadena
                                                                 
                                                                 Caracter = Caracter + CaracteresConcepto
                                                                 
                                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA/////////////////
                                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                                 If Len(cadena) < CaracteresConcepto Then
                                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                                     LineaConcepto = LineaConcepto + 1
                                                                     Printer.CurrentX = Val(X18) '25
                                                                     Printer.CurrentY = Val(Y18) + (5 * LineaConcepto)
                                                                     Printer.FontName = "Times New Roman"
                                                                     Printer.FontSize = 11
                                                                     Printer.FontBold = True
                                                                     Printer.Print cadena
                                                                     
                                                                     Caracter = Caracter + CaracteresConcepto
                                                                 End If
                                                    
                                                 End If
                                          Loop
                                          
                                     Else
                                                    Printer.CurrentX = Val(X18) '5
                                                    Printer.CurrentY = Val(Y18) + (5 * i) '120
                                                    Printer.FontName = "Times New Roman"
                                                    Printer.FontSize = 11
                                                    Printer.FontBold = True
                                                    Printer.Print Concepto
                                     End If
                                   End If
                                    
                                    Dia = Day(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                    mes = Month(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                    Año = Year(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                    Meses = Month(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                  
                                '    FrmCheque.DtaConsulta.Recordset.MoveLast
                                   If X9 <> 0 Or Y9 <> 0 Then
                                    Printer.CurrentX = X9
                                    Printer.CurrentY = Y9
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 11
                                    Printer.FontBold = True
                '                    Printer.Print FrmCheque.DtaConsulta.Recordset("NumeroMovimiento")
                                    Printer.Print Me.LblConsecutivo.Text
                                   End If
                                    
                                   If X1 <> 0 Or Y1 <> 0 Then
                                    Printer.CurrentX = Val(X1)
                                    Printer.CurrentY = Val(Y1) + (5 * i)
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 9
                                    Printer.FontBold = True
                                    Printer.Print FrmCheque.TxtNombre.Text
                                   End If
                                   
                                  If X14 <> 0 Or Y14 <> 0 Then
                                    Printer.CurrentX = Val(X14)
                                    Printer.CurrentY = Val(Y14) + (5 * i)
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 9
                                    Printer.FontBold = True
                                    Printer.Print FrmCheque.TxtNombre.Text
                                   End If
                                   
                                   If X4 <> 0 Or Y4 <> 0 Then
                                    Printer.CurrentX = Val(X4)
                                    Printer.CurrentY = Val(Y4) + (5 * i)
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 9
                                    Printer.FontBold = True
                                    Printer.Print FrmCheque.TxtLetras.Text
                                   End If
                                   
                                  If X15 <> 0 Or Y15 <> 0 Then
                                    Printer.CurrentX = Val(X15)
                                    Printer.CurrentY = Val(Y15) + (5 * i)
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 9
                                    Printer.FontBold = True
                                    Printer.Print FrmCheque.TxtLetras.Text
                                   End If
                                   
                                   If X3 <> 0 Or Y3 <> 0 Then
                                    Printer.CurrentX = Val(X3)
                                    Printer.CurrentY = Val(Y3) + (5 * i)
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 11
                                    Printer.FontBold = True
                                    Printer.Print Format(Monto, "##,##0.00")
                                   End If
                                   
                                                      
                                   If X16 <> 0 Or Y16 <> 0 Then
                                    Printer.CurrentX = Val(X16)
                                    Printer.CurrentY = Val(Y16) + (5 * i)
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 11
                                    Printer.FontBold = True
                                    Printer.Print Format(Monto, "##,##0.00")
                                   End If
                                
                                   If X2 <> 0 Or Y2 <> 0 Then
                                    Printer.CurrentX = Val(X2) '20
                                    Printer.CurrentY = Val(Y2) '288
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 11
                                    Printer.FontBold = True
                                    FechaLetra = "Managua          " & Format(Day(FrmCheque.DtaConsulta.Recordset("FechaTransaccion")), "0#") & "          " & Format(Month(FrmCheque.DtaConsulta.Recordset("FechaTransaccion")), "0#") & "           " & Year(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                    Printer.Print FechaLetra
                                   End If
                                
                                    If X17 <> 0 Or Y17 <> 0 Then
                                    Printer.CurrentX = Val(X17) '20
                                    Printer.CurrentY = Val(Y17) '288
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 11
                                    Printer.FontBold = True
                                    FechaLetra = "Managua          " & Format(Day(FrmCheque.DtaConsulta.Recordset("FechaTransaccion")), "0#") & "          " & Format(Month(FrmCheque.DtaConsulta.Recordset("FechaTransaccion")), "0#") & "           " & Year(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                                    Printer.Print FechaLetra
                                   End If
                                
                   End If

        
        Printer.EndDoc


Else

 FrmCheque.AdoCordenadas.RecordSource = "SELECT CodCuenta, X1, Y1, X2, Y2, X3, Y3, X4, Y4, X5, Y5, X6, Y6, X7, Y7, X8, Y8, X9, Y9, X10, Y10, X11, Y11, X12, Y12, X13, Y13,X14, Y14,X15, Y15,X16, Y16,X17, Y17, X18, Y18, X19, Y19,X20, Y20,X21, Y21, X22, Y22, NLineas,CaracteresLineas, CaracteresConcepto From CordenadasCheque WHERE  (CodCuenta = '" & CodigoCuenta & "')"
 FrmCheque.AdoCordenadas.Refresh
 If FrmCheque.AdoCordenadas.Recordset.EOF Then
  MsgBox "No Existen Coordenadas, para la Cuenta", vbCritical, "Sistema Contable"
  Exit Sub
 End If
 
 
        X1 = FrmCheque.AdoCordenadas.Recordset("X1")
        Y1 = FrmCheque.AdoCordenadas.Recordset("Y1")
        X2 = FrmCheque.AdoCordenadas.Recordset("X2")
        Y2 = FrmCheque.AdoCordenadas.Recordset("Y2")
        X3 = FrmCheque.AdoCordenadas.Recordset("X3")
        Y3 = FrmCheque.AdoCordenadas.Recordset("Y3")
        X4 = FrmCheque.AdoCordenadas.Recordset("X4")
        Y4 = FrmCheque.AdoCordenadas.Recordset("Y4")
        X5 = FrmCheque.AdoCordenadas.Recordset("X5")
        Y5 = FrmCheque.AdoCordenadas.Recordset("Y5")
        X6 = FrmCheque.AdoCordenadas.Recordset("X6")
        Y6 = FrmCheque.AdoCordenadas.Recordset("Y6")
        X7 = FrmCheque.AdoCordenadas.Recordset("X7")
        Y7 = FrmCheque.AdoCordenadas.Recordset("Y7")
        X8 = FrmCheque.AdoCordenadas.Recordset("X8")
        Y8 = FrmCheque.AdoCordenadas.Recordset("Y8")
        X9 = FrmCheque.AdoCordenadas.Recordset("X9")
        Y9 = FrmCheque.AdoCordenadas.Recordset("Y9")
        X10 = FrmCheque.AdoCordenadas.Recordset("X10")
        Y10 = FrmCheque.AdoCordenadas.Recordset("Y10")
        X11 = FrmCheque.AdoCordenadas.Recordset("X11")
        Y11 = FrmCheque.AdoCordenadas.Recordset("Y11")
        X12 = FrmCheque.AdoCordenadas.Recordset("X12")
        Y12 = FrmCheque.AdoCordenadas.Recordset("Y12")
        X13 = FrmCheque.AdoCordenadas.Recordset("X13")
        Y13 = FrmCheque.AdoCordenadas.Recordset("Y13")
        X14 = FrmCheque.AdoCordenadas.Recordset("X14")
        Y14 = FrmCheque.AdoCordenadas.Recordset("Y14")
        X15 = FrmCheque.AdoCordenadas.Recordset("X15")
        Y15 = FrmCheque.AdoCordenadas.Recordset("Y15")
        X16 = FrmCheque.AdoCordenadas.Recordset("X16")
        Y16 = FrmCheque.AdoCordenadas.Recordset("Y16")
        X17 = FrmCheque.AdoCordenadas.Recordset("X17")
        Y17 = FrmCheque.AdoCordenadas.Recordset("Y17")
        X18 = FrmCheque.AdoCordenadas.Recordset("X18")
        Y18 = FrmCheque.AdoCordenadas.Recordset("Y18")
        X19 = FrmCheque.AdoCordenadas.Recordset("X19")
        Y19 = FrmCheque.AdoCordenadas.Recordset("Y19")
        X20 = FrmCheque.AdoCordenadas.Recordset("X20")
        Y20 = FrmCheque.AdoCordenadas.Recordset("Y20")
        X21 = FrmCheque.AdoCordenadas.Recordset("X21")
        Y21 = FrmCheque.AdoCordenadas.Recordset("Y21")
        X22 = FrmCheque.AdoCordenadas.Recordset("X22")
        Y22 = FrmCheque.AdoCordenadas.Recordset("Y22")
        NLineas = Val(FrmCheque.AdoCordenadas.Recordset("NLineas"))
        CaracteresLineas = Val(FrmCheque.AdoCordenadas.Recordset("CaracteresLineas"))
        CaracteresConcepto = Val(FrmCheque.AdoCordenadas.Recordset("CaracteresConcepto"))
        
         If FrmCheque.CmbMoneda.Text = "Córdobas" Then
            If FrmCheque.ChkCheque.Value = 1 Then
              TasaCambio = BuscaTasaCambio(FrmCheque.TxtFecha.Value)
              Monto = FrmCheque.TxtMonto.Text
              Monto = Monto / TasaCambio
              FrmCheque.TxtLetras.Text = sw.ConvertCurrencyToSpanish(Monto, "Dólares")
'              ArepCheque.LblMonto.Caption = Format(Monto, "##,##0.00")
            Else
              Monto = FrmCheque.TxtMonto.Text
            End If
         Else
           Monto = FrmCheque.TxtMonto.Text
         End If
 
'Cargo la Consulta del Cheque
 FrmCheque.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento,Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito,Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas,Transacciones.NumeroMovimiento , Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas1 & "', 102))AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ")ORDER BY Transacciones.NTransaccion"
 FrmCheque.DtaConsulta.Refresh
 Printer.FontSize = 9
 'Inicio el Ciclo de Impresion
 i = 1
 
 TasaCambio = BuscaTasaCambio(FrmCheque.TxtFecha.Value)
 Do While Not FrmCheque.DtaConsulta.Recordset.EOF
   
   '////////////////////////////////////////////////////////////////////////////////////////////
   '//////////////////////SI EL NUMERO DE LINEAS ES MAYOR CREO UNA NUEVA PAGINA////////////////
   '///////////////////////////////////////////////////////////////////////////////////////////
    
    If i > NLineas Then
      Printer.NewPage
      i = 1
    End If
    
          
            If i = 1 Then
        ContadorLinea = i
          
                  '//////////////////////////////////////////////////////////////////////////////////////
                  '//////////////////////////IMPRIMO LOS ENCABEZADOS ////////////////////////////////////
                  '//////////////////////////////////////////////////////////////////////////////////
           
                   If X5 <> 0 Or Y5 <> 0 Then
                     Caracter = 1
                     LineaConcepto = 1
                     cadena = Concepto
                     If Len(cadena) > CaracteresConcepto Then
                          Do While Len(cadena) >= CaracteresConcepto
                                 If Caracter = 1 Then
'                                    Printer.CurrentX = Val(X5) '5
'                                    Printer.CurrentY = Val(Y5) + (5 * i) '120
'                                    Printer.FontName = "Times New Roman"
'                                    Printer.FontSize = 11
'                                    Printer.FontBold = True
'                                    Printer.Print Concepto
                                    
                                           
                                                 cadena = Mid(Concepto, 1, CaracteresConcepto)
                                                 Printer.CurrentX = Val(X5) '25
                                                 Printer.CurrentY = Val(Y5) + (5 * LineaConcepto)
                                                 Printer.FontName = "Times New Roman"
                                                 Printer.FontSize = 11
                                                 Printer.FontBold = True
                                                 Printer.Print cadena
                                                 Caracter = Caracter + CaracteresConcepto
                                                 
                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA SIGUIENTE/////////////////
                                                 
                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                 If Len(cadena) < CaracteresConcepto Then
                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                     LineaConcepto = LineaConcepto + 1
                                                     Printer.CurrentX = Val(X5) '25
                                                     Printer.CurrentY = Val(Y5) + (5 * LineaConcepto)
                                                     Printer.FontName = "Times New Roman"
                                                     Printer.FontSize = 11
                                                     Printer.FontBold = True
                                                     Printer.Print cadena
                                                     
                                                     Caracter = Caracter + CaracteresConcepto
                                                 End If
                                                 
                                 Else
                                                 
                                                 LineaConcepto = LineaConcepto + 1
                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                 Printer.CurrentX = Val(X5) '25
                                                 Printer.CurrentY = Val(Y5) + (5 * LineaConcepto)
                                                 Printer.FontName = "Times New Roman"
                                                 Printer.FontSize = 11
                                                 Printer.FontBold = True
                                                 Printer.Print cadena
                                                 
                                                 Caracter = Caracter + CaracteresConcepto
                                                 
                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA/////////////////
                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                 If Len(cadena) < CaracteresConcepto Then
                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                     LineaConcepto = LineaConcepto + 1
                                                     Printer.CurrentX = Val(X5) '25
                                                     Printer.CurrentY = Val(Y5) + (5 * LineaConcepto)
                                                     Printer.FontName = "Times New Roman"
                                                     Printer.FontSize = 11
                                                     Printer.FontBold = True
                                                     Printer.Print cadena
                                                     
                                                     Caracter = Caracter + CaracteresConcepto
                                                 End If
                                    
                                 End If
                          Loop
                          
                     Else
                                    Printer.CurrentX = Val(X5) '5
                                    Printer.CurrentY = Val(Y5) + (5 * i) '120
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 11
                                    Printer.FontBold = True
                                    Printer.Print Concepto
                     End If
                   End If
                   
                   
                  If X18 <> 0 Or Y18 <> 0 Then
                     Caracter = 1
                     LineaConcepto = 1
                     cadena = Concepto
                     If Len(cadena) > CaracteresConcepto Then
                          Do While Len(cadena) >= CaracteresConcepto
                                 If Caracter = 1 Then
'                                    Printer.CurrentX = Val(X5) '5
'                                    Printer.CurrentY = Val(Y5) + (5 * i) '120
'                                    Printer.FontName = "Times New Roman"
'                                    Printer.FontSize = 11
'                                    Printer.FontBold = True
'                                    Printer.Print Concepto
                                    
                                           
                                                 cadena = Mid(Concepto, 1, CaracteresConcepto)
                                                 Printer.CurrentX = Val(X18) '25
                                                 Printer.CurrentY = Val(Y18) + (5 * LineaConcepto)
                                                 Printer.FontName = "Times New Roman"
                                                 Printer.FontSize = 11
                                                 Printer.FontBold = True
                                                 Printer.Print cadena
                                                 Caracter = Caracter + CaracteresConcepto
                                                 
                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA SIGUIENTE/////////////////
                                                 
                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                 If Len(cadena) < CaracteresConcepto Then
                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                     LineaConcepto = LineaConcepto + 1
                                                     Printer.CurrentX = Val(X18) '25
                                                     Printer.CurrentY = Val(Y18) + (5 * LineaConcepto)
                                                     Printer.FontName = "Times New Roman"
                                                     Printer.FontSize = 11
                                                     Printer.FontBold = True
                                                     Printer.Print cadena
                                                     
                                                     Caracter = Caracter + CaracteresConcepto
                                                 End If
                                                 
                                 Else
                                                 
                                                 LineaConcepto = LineaConcepto + 1
                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                 Printer.CurrentX = Val(X18) '25
                                                 Printer.CurrentY = Val(Y18) + (5 * LineaConcepto)
                                                 Printer.FontName = "Times New Roman"
                                                 Printer.FontSize = 11
                                                 Printer.FontBold = True
                                                 Printer.Print cadena
                                                 
                                                 Caracter = Caracter + CaracteresConcepto
                                                 
                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA/////////////////
                                                 cadena = Mid(Concepto, Caracter, CaracteresConcepto)
                                                 If Len(cadena) < CaracteresConcepto Then
                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                     LineaConcepto = LineaConcepto + 1
                                                     Printer.CurrentX = Val(X18) '25
                                                     Printer.CurrentY = Val(Y18) + (5 * LineaConcepto)
                                                     Printer.FontName = "Times New Roman"
                                                     Printer.FontSize = 11
                                                     Printer.FontBold = True
                                                     Printer.Print cadena
                                                     
                                                     Caracter = Caracter + CaracteresConcepto
                                                 End If
                                    
                                 End If
                          Loop
                          
                     Else
                                    Printer.CurrentX = Val(X18) '5
                                    Printer.CurrentY = Val(Y18) + (5 * i) '120
                                    Printer.FontName = "Times New Roman"
                                    Printer.FontSize = 11
                                    Printer.FontBold = True
                                    Printer.Print Concepto
                     End If
                   End If
                    
                    Dia = Day(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                    mes = Month(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                    Año = Year(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                    Meses = Month(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                  
                '    FrmCheque.DtaConsulta.Recordset.MoveLast
                   If X9 <> 0 Or Y9 <> 0 Then
                    Printer.CurrentX = X9
                    Printer.CurrentY = Y9
                    Printer.FontName = "Times New Roman"
                    Printer.FontSize = 11
                    Printer.FontBold = True
'                    Printer.Print FrmCheque.DtaConsulta.Recordset("NumeroMovimiento")
                    Printer.Print Me.LblConsecutivo.Text
                   End If
                    
                   If X1 <> 0 Or Y1 <> 0 Then
                    Printer.CurrentX = Val(X1)
                    Printer.CurrentY = Val(Y1) + (5 * i)
                    Printer.FontName = "Times New Roman"
                    Printer.FontSize = 11
                    Printer.FontBold = True
                    Printer.Print FrmCheque.TxtNombre.Text
                   End If
                   
                  If X14 <> 0 Or Y14 <> 0 Then
                    Printer.CurrentX = Val(X14)
                    Printer.CurrentY = Val(Y14) + (5 * i)
                    Printer.FontName = "Times New Roman"
                    Printer.FontSize = 11
                    Printer.FontBold = True
                    Printer.Print FrmCheque.TxtNombre.Text
                   End If
                   
                   If X4 <> 0 Or Y4 <> 0 Then
                    Printer.CurrentX = Val(X4)
                    Printer.CurrentY = Val(Y4) + (5 * i)
                    Printer.FontName = "Times New Roman"
                    Printer.FontSize = 9
                    Printer.FontBold = True
                    Printer.Print FrmCheque.TxtLetras.Text
                   End If
                   
                  If X15 <> 0 Or Y15 <> 0 Then
                    Printer.CurrentX = Val(X15)
                    Printer.CurrentY = Val(Y15) + (5 * i)
                    Printer.FontName = "Times New Roman"
                    Printer.FontSize = 9
                    Printer.FontBold = True
                    Printer.Print FrmCheque.TxtLetras.Text
                   End If
                   
                   If X3 <> 0 Or Y3 <> 0 Then
                    Printer.CurrentX = Val(X3)
                    Printer.CurrentY = Val(Y3) + (5 * i)
                    Printer.FontName = "Times New Roman"
                    Printer.FontSize = 11
                    Printer.FontBold = True
                    Printer.Print Format(Monto, "##,##0.00")
                   End If
                   
                                      
                   If X16 <> 0 Or Y16 <> 0 Then
                    Printer.CurrentX = Val(X16)
                    Printer.CurrentY = Val(Y16) + (5 * i)
                    Printer.FontName = "Times New Roman"
                    Printer.FontSize = 11
                    Printer.FontBold = True
                    Printer.Print Format(Monto, "##,##0.00")
                   End If
                
                   If X2 <> 0 Or Y2 <> 0 Then
                    Printer.CurrentX = Val(X2) '20
                    Printer.CurrentY = Val(Y2) '288
                    Printer.FontName = "Times New Roman"
                    Printer.FontSize = 11
                    Printer.FontBold = True
                    FechaLetra = "Managua          " & Format(Day(FrmCheque.DtaConsulta.Recordset("FechaTransaccion")), "0#") & "          " & Format(Month(FrmCheque.DtaConsulta.Recordset("FechaTransaccion")), "0#") & "           " & Year(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                    Printer.Print FechaLetra
                   End If
                
                    If X17 <> 0 Or Y17 <> 0 Then
                    Printer.CurrentX = Val(X17) '20
                    Printer.CurrentY = Val(Y17) '288
                    Printer.FontName = "Times New Roman"
                    Printer.FontSize = 11
                    Printer.FontBold = True
                    FechaLetra = "Managua          " & Format(Day(FrmCheque.DtaConsulta.Recordset("FechaTransaccion")), "0#") & "          " & Format(Month(FrmCheque.DtaConsulta.Recordset("FechaTransaccion")), "0#") & "           " & Year(FrmCheque.DtaConsulta.Recordset("FechaTransaccion"))
                    Printer.Print FechaLetra
                   End If
                
           End If
           
           '//////////////////////////////////////////////////////////////////////////////////////
           '//////////////////////////IMPRIMO LOS DETALLES ////////////////////////////////////
           '//////////////////////////////////////////////////////////////////////////////////
           
           
           If X6 <> 0 Or Y6 <> 0 Then
            Printer.CurrentX = Val(X6) '5
            Printer.CurrentY = Val(Y6) + (5 * i)
            cadena = FrmCheque.DtaConsulta.Recordset("CodCuentas")
            If Len(cadena) > 20 Then
             cadena = Mid(cadena, 1, 20)
            End If
            
            Printer.FontName = "Times New Roman"
            Printer.FontSize = 9
            Printer.FontBold = False
            Printer.Print cadena
           End If
        
        
        
        
          If X10 <> 0 Or Y10 <> 0 Then
            Printer.CurrentX = Val(X10) '25
            Printer.CurrentY = Val(Y10) + (5 * i)
            cadena = FrmCheque.DtaConsulta.Recordset("NombreCuenta")
            If Len(cadena) > 24 Then
             cadena = Mid(cadena, 1, 24)
            End If
            
            Printer.FontName = "Times New Roman"
            Printer.FontSize = 9
            Printer.FontBold = False
            Printer.Print cadena
          End If
        
         
            If X11 <> 0 Or Y11 <> 0 Then
                     CadenaDescripcion = FrmCheque.DtaConsulta.Recordset("DescripcionMovimiento")
                     cadena = FrmCheque.DtaConsulta.Recordset("DescripcionMovimiento")
                     Caracter = 1
                     ContadorLinea = i
                     
                     If Len(cadena) > CaracteresLineas Then
                              Do While Len(cadena) >= CaracteresLineas
                                       If Caracter = 1 Then
                                                 cadena = Mid(cadena, 1, CaracteresLineas)
                                                 Printer.CurrentX = Val(X11) '25
                                                 Printer.CurrentY = Val(Y11) + (5 * i)
                                                 Printer.FontName = "Times New Roman"
                                                 Printer.FontSize = 9
                                                 Printer.FontBold = False
                                                 Printer.Print cadena
                                                 Caracter = Caracter + CaracteresLineas
                                                 
                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA SIGUIENTE/////////////////
                                                 cadena = Mid(CadenaDescripcion, Caracter, CaracteresLineas)
                                                 If Len(cadena) < CaracteresLineas Then
                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                     ContadorLinea = ContadorLinea + 1
                                                     Printer.CurrentX = Val(X11) '25
                                                     Printer.CurrentY = Val(Y11) + (5 * ContadorLinea)
                                                     Printer.FontName = "Times New Roman"
                                                     Printer.FontSize = 9
                                                     Printer.FontBold = False
                                                     Printer.Print cadena
                                                     
                                                     Caracter = Caracter + CaracteresLineas
                                                 End If
                                         Else
                                                 ContadorLinea = ContadorLinea + 1
                                                 cadena = Mid(CadenaDescripcion, Caracter, CaracteresLineas)
                                                 Printer.CurrentX = Val(X11) '25
                                                 Printer.CurrentY = Val(Y11) + (5 * ContadorLinea)
                                                 Printer.FontName = "Times New Roman"
                                                 Printer.FontSize = 9
                                                 Printer.FontBold = False
                                                 Printer.Print cadena
                                                 
                                                 Caracter = Caracter + CaracteresLineas
                                                 
                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA/////////////////
                                                 cadena = Mid(CadenaDescripcion, Caracter, CaracteresLineas)
                                                 If Len(cadena) < CaracteresLineas Then
                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                     ContadorLinea = ContadorLinea + 1
                                                     Printer.CurrentX = Val(X11) '25
                                                     Printer.CurrentY = Val(Y11) + (5 * ContadorLinea)
                                                     Printer.FontName = "Times New Roman"
                                                     Printer.FontSize = 9
                                                     Printer.FontBold = False
                                                     Printer.Print cadena
                                                     
                                                     Caracter = Caracter + CaracteresLineas
                                                 End If
                                                 
                                                 
                                        End If
                 
                                              
                              Loop
                                              
                     Else
                             Printer.CurrentX = Val(X11) '25
                             Printer.CurrentY = Val(Y11) + (5 * i)
                             Printer.FontName = "Times New Roman"
                             Printer.FontSize = 9
                             Printer.FontBold = False
                             Printer.Print cadena
                                           
                     End If
                  

            End If
   
          If FrmCheque.CmbMoneda.Text = "Córdobas" Then
             DebitoCordobas = CDbl(FrmCheque.DtaConsulta.Recordset("Debito"))
             DebitoDolares = DebitoCordobas / TasaCambio
             
             CreditoCordobas = FrmCheque.DtaConsulta.Recordset("Credito")
             CreditoDolares = CreditoCordobas / TasaCambio
           
          Else
             DebitoDolares = FrmCheque.DtaConsulta.Recordset("Debito")
             DebitoCordobas = DebitoDolares * TasaCambio
             
             CreditoDolares = FrmCheque.DtaConsulta.Recordset("Credito")
             CreditoCordobas = CreditoCordobas * TasaCambio
             
        
          End If
              
        
          
           If X12 <> 0 Or Y12 <> 0 Then
            Printer.CurrentX = Val(X12) '135
             Printer.CurrentY = Val(Y12) + (5 * i)
             Printer.FontName = "Times New Roman"
             Printer.FontSize = 9
             Printer.FontBold = False
             Printer.Print Format(DebitoCordobas, "##,##0.00")
'            Printer.Print Format(FrmCheque.DtaConsulta.Recordset("Debito"), "##,##0.00")
           End If
           
          If X19 <> 0 Or Y19 <> 0 Then
            Printer.CurrentX = Val(X19) '135
             Printer.CurrentY = Val(Y19) + (5 * i)
             Printer.FontName = "Times New Roman"
             Printer.FontSize = 9
             Printer.FontBold = False
             Printer.Print Format(DebitoDolares, "##,##0.00")
'            Printer.Print Format(FrmCheque.DtaConsulta.Recordset("Debito"), "##,##0.00")
           End If
        
        
           If X13 <> 0 Or Y13 <> 0 Then
            Printer.CurrentX = Val(X13) '165
              Printer.CurrentY = Val(Y13) + (5 * i) '165
            Printer.FontName = "Times New Roman"
            Printer.FontSize = 9
            Printer.FontBold = False
            Printer.Print Format(CreditoCordobas, "##,##0.00")
'            Printer.Print Format(FrmCheque.DtaConsulta.Recordset("Credito"), "##,##0.00")
           End If
           
           If X20 <> 0 Or Y20 <> 0 Then
            Printer.CurrentX = Val(X20) '165
              Printer.CurrentY = Val(Y20) + (5 * i) '165
            Printer.FontName = "Times New Roman"
            Printer.FontSize = 9
            Printer.FontBold = False
            Printer.Print Format(CreditoDolares, "##,##0.00")
'            Printer.Print Format(FrmCheque.DtaConsulta.Recordset("Credito"), "##,##0.00")
           End If
           
           
            If i > 1 Then
              UltimaLinea = UltimaLinea + (5 * i) + DiferenciaY - 4
            End If
           
            i = ContadorLinea
            i = i + 1
            ContadorLinea = i
        '
        
        ' 'Fin del Ciclo


 FrmCheque.DtaConsulta.Recordset.MoveNext
 Loop
'  i = 4
'     i = 4
'    Printer.CurrentX = 70
'    Printer.CurrentY = 140 + (5 * i)
'    Printer.Print "Total"


          If FrmCheque.CmbMoneda.Text = "Córdobas" Then
             TotalDebitoCordobas = TotalDebito
             TotalDebitoDolares = TotalDebitoCordobas / TasaCambio
             
             TotalCreditoCordobas = TotalCredito
             TotalCreditoDolares = TotalCreditoCordobas / TasaCambio
           
          Else
             TotalDebitoDolares = TotalDebito
             TotalDebitoCordobas = TotalDebitoCordobas * TasaCambio
             
             TotalCreditoDolares = TotalCredito
             TotalCreditoCordobas = TotalCreditoDolares * TasaCambio
             
        
          End If

  If X7 <> 0 Or Y7 <> 0 Then
    Printer.CurrentX = Val(X7) '135
    Printer.CurrentY = Val(Y7) '288
    Printer.FontName = "Times New Roman"
    Printer.FontSize = 9
    Printer.FontBold = False
    Printer.Print Format(TotalDebitoCordobas, "##,##0.00")
  End If
  
 If X21 <> 0 Or Y21 <> 0 Then
    Printer.CurrentX = Val(X21) '135
    Printer.CurrentY = Val(Y21) '288
    Printer.FontName = "Times New Roman"
    Printer.FontSize = 9
    Printer.FontBold = False
    Printer.Print Format(TotalDebitoDolares, "##,##0.00")
  End If
  
  If X8 <> 0 Or Y8 <> 0 Then
    Printer.CurrentX = Val(X8) '165
    Printer.CurrentY = Val(Y8) '288
    Printer.FontName = "Times New Roman"
    Printer.FontSize = 9
    Printer.FontBold = False
    Printer.Print Format(TotalCreditoCordobas, "##,##0.00")
  End If
  
 If X22 <> 0 Or Y22 <> 0 Then
    Printer.CurrentX = Val(X22) '165
    Printer.CurrentY = Val(Y22) '288
    Printer.FontName = "Times New Roman"
    Printer.FontSize = 9
    Printer.FontBold = False
    Printer.Print Format(TotalCreditoDolares, "##,##0.00")
  End If
'
'    Printer.CurrentX = Val(X2) '20
'    Printer.CurrentY = Val(Y2) '288
'    Printer.Print Dia
'
'    Printer.CurrentX = Val(X2) + 18 '38
'    Printer.CurrentY = Y2 '288
'    Printer.Print Mes
'
'    Printer.CurrentX = Val(X2) + 38 '58
'    Printer.CurrentY = Y2 '288
'    Printer.Print Año


'    Printer.CurrentX = 100
'    Printer.CurrentY = 198 + (5 * i)
'    Printer.Print Fechass
'    Printer.CurrentX = 19
'    Printer.CurrentY = 206 + (5 * i)
'    Printer.Print FrmCheque.TxtNombre.Text
'    Printer.CurrentX = 4
'    Printer.CurrentY = 215 + (5 * i)
'    Printer.Print FrmCheque.TxtLetras.Text
'    Printer.CurrentX = 140
'    Printer.CurrentY = 206 + (5 * i)
'    Printer.Print Format(FrmCheque.TxtMonto.Text, "##,##0.00")

   

 
'termino de imprimir las facturas
Printer.EndDoc



















' ArepCheque.DtaCheque.ConnectionString = ConexionReporte
' ArepCheque.LblMemo = FrmCheque.TxtMemo
' ArepCheque.LblNombre2.Caption = FrmCheque.TxtNombre.Text
' ArepCheque.LblChequeNo.Caption = Me.LblConsecutivo.Text
' ArepCheque.Field15.Visible = False
' ArepCheque.DtaCheque.Source = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento,Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito,Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas,Transacciones.NumeroMovimiento , Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas1 & "', 102))AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ")ORDER BY Transacciones.NTransaccion"
' ArepCheque.Show 1

End If


'Unload Me


End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd
End Sub

Private Sub Form_Initialize()
On Error GoTo TipoErrs
Dim SqlCheque As String
    Set ew = New cls_NumEnglishWord
    Set sw = New cls_NumSpanishWord
    'DBGdetalleCk.Columns(3).Button = True
Exit Sub
TipoErrs:
ControlErrores
End Sub
