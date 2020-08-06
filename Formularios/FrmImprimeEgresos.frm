VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmImprimeEgresos 
   Caption         =   "Imprime Egreso"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   2715
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton SmartButton1 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Encabezados"
      Height          =   1095
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir Comprobante"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Consecutivo"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2535
      Begin VB.CheckBox ChkOcultar 
         Caption         =   "Ocultar Detalle"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox LblConsecutivo 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmImprimeEgresos.frx":0000
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblCuenta 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "FrmImprimeEgresos.frx":0072
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "FrmImprimeEgresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ew As cls_NumEnglishWord
Private sw As cls_NumSpanishWord

Private Sub CmdGrabar_Click()
Unload Me
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


Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd
End Sub

Private Sub SmartButton1_Click()
Dim Fechas1 As String, Fechas2 As String
Dim CodigoCuenta As String, Concepto As String
Dim x, y, H, V, Page As Integer, Dia As String, mes As String, Año As String
Dim i, J As Integer, Fechass As Date
Dim TotalDebito, TotalCredito, Totalpag As Double
Dim SubTotal, Total, IGV As Double, Cadena As String
Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double, X5 As Double, Y5 As Double, X6 As Double, Y6 As Double, X7 As Double, Y7 As Double, X8 As Double, Y8 As Double, X9 As Double, Y9 As Double, X10 As Double, Y10 As Double, X11 As Double, Y11 As Double, X12 As Double, Y12 As Double, X13 As Double, Y13 As Double
Dim UltimaLinea As Double, DiferenciaY As Double, NLineas As Double
Dim Caracter As Double, ContadorLinea As Double, CadenaDescripcion As String, CaracteresLineas As Double
Dim Meses As Double, Monto As Double, TasaCambio As Double


Page = 1

Printer.FontSize = 6
Printer.ScaleMode = 6



TotalDebito = 0
TotalCredito = 0

If Not FrmEgresos.TxtMemo.Text = "" Then
 Concepto = FrmEgresos.TxtMemo.Text
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
      NumFecha1 = FrmEgresos.TxtFecha.Value
      Fechas1 = Format(FrmEgresos.TxtFecha.Value, "YYYY/MM/DD")
      NMovimiento = Val(FrmEgresos.TxtNTransacciones)
      FrmEgresos.DtaConsulta.RecordSource = "SELECT     FechaTransaccion, CodCuentas, NTransaccion, NumeroMovimiento, TCambio * Debito AS MDebito, TCambio * Credito AS MCredito, TCambio, Debito, Credito From Transacciones WHERE     (FechaTransaccion = CONVERT(DATETIME, '" & Fechas1 & "', 102)) AND (NumeroMovimiento = " & NMovimiento & ")"
'      FrmEgresos.DtaConsulta.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.CodCuentas, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, [Transacciones]![TCambio]*[Transacciones]![Debito] AS MDebito, [Transacciones]![TCambio]*[Transacciones]![Credito] AS MCredito, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito From Transacciones WHERE (((Transacciones.FechaTransaccion)=" & NumFecha1 & ") AND ((Transacciones.NumeroMovimiento)=" & NMovimiento & "))"
      FrmEgresos.DtaConsulta.Refresh
      Do While Not FrmEgresos.DtaConsulta.Recordset.EOF
      If FrmEgresos.TxtMonto.Text = "" Then
       MontoCheque = 0
      Else
       MontoCheque = FrmEgresos.TxtMonto
      End If
       Debito = FrmEgresos.DtaConsulta.Recordset("Credito")
       Credito = FrmEgresos.DtaConsulta.Recordset("Credito")
       TotalDebito = TotalDebito + Debito
       TotalCredito = TotalCredito + Credito
       FrmEgresos.DtaConsulta.Recordset.MoveNext
      Loop
      
  CodigoCuenta = FrmEgresos.DBCodigo.Text
  FrmEgresos.DtaConsulta.RecordSource = "SELECT CodCuentas, ConsecutivoCheque From NConsecutivos WHERE (CodCuentas = '" & CodigoCuenta & "')"
  FrmEgresos.DtaConsulta.Refresh
  If Not FrmEgresos.DtaConsulta.Recordset.EOF Then
      FrmEgresos.DtaConsulta.Recordset("ConsecutivoCheque") = Me.LblConsecutivo.Text
      FrmEgresos.DtaConsulta.Recordset.Update
  End If
  
  
  FrmEgresos.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento,Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito,Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas,Transacciones.NumeroMovimiento , Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas1 & "', 102))AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ")ORDER BY Transacciones.NTransaccion"
  FrmEgresos.DtaConsulta.Refresh
  Do While Not FrmEgresos.DtaConsulta.Recordset.EOF
  
    FrmEgresos.DtaConsulta.Recordset("ChequeNo") = Me.LblConsecutivo.Text
    FrmEgresos.DtaConsulta.Recordset.Update
    FrmEgresos.DtaConsulta.Recordset.MoveNext
  Loop
  
  Me.LblConsecutivo.Text = Format(Val(Me.LblConsecutivo.Text), "00##")
  
If Me.Check1.Value = 1 Then

 ArepEgresos.DtaCheque.ConnectionString = ConexionReporte
 
 If FrmEgresos.CmbMoneda.Text = "Córdobas" Then
    ArepEgresos.LblMontoSimbolo.Caption = "Monto C$"
   If FrmEgresos.ChkCheque.Value = 1 Then
     TasaCambio = BuscaTasaCambio(FrmEgresos.TxtFecha.Value)
     Monto = FrmEgresos.TxtMonto.Text
     Monto = Monto / TasaCambio
     ArepEgresos.LblDescripcionMonto.Caption = sw.ConvertCurrencyToSpanish(Monto, "Dólares")
     ArepEgresos.LblMonto.Caption = Format(Monto, "##,##0.00")
   Else
    ArepEgresos.LblDescripcionMonto.Caption = FrmEgresos.TxtLetras.Text
    ArepEgresos.LblMonto.Caption = Format(FrmEgresos.TxtMonto.Text, "##,##0.00")
   End If
 Else
    ArepEgresos.LblMontoSimbolo.Caption = "Monto $"
    ArepEgresos.LblDescripcionMonto.Caption = FrmEgresos.TxtLetras.Text
    ArepEgresos.LblMonto.Caption = Format(FrmEgresos.TxtMonto.Text, "##,##0.00")
 End If
 ArepEgresos.LblMemo.Caption = FrmEgresos.TxtMemo.Text
 
 ArepEgresos.LblNombre.Caption = FrmEgresos.TxtNombre.Text
 ArepEgresos.LblChequeNo.Caption = Me.LblConsecutivo.Text
 If Me.ChkOcultar.Value = 1 Then
   ArepEgresos.Detail.Visible = False
 Else
   ArepEgresos.Detail.Visible = True
 End If
' ArepEgresos.LblEmpresa = FrmEgresos.DtaDatosEmpresa.Recordset("NombreEmpresa")
' ArepEgresos.LblEmpresa1 = FrmEgresos.DtaDatosEmpresa.Recordset("Direccion")
 ArepEgresos.DtaCheque.Source = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento,Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito,Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas,Transacciones.NumeroMovimiento , Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas1 & "', 102))AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ")ORDER BY Transacciones.NTransaccion"
 ArepEgresos.Show 1
Else

 FrmEgresos.AdoCordenadas.RecordSource = "SELECT CodCuenta, X1, Y1, X2, Y2, X3, Y3, X4, Y4, X5, Y5, X6, Y6, X7, Y7, X8, Y8, X9, Y9, X10, Y10, X11, Y11, X12, Y12, X13, Y13, NLineas,CaracteresLineas From CordenadasCheque WHERE  (CodCuenta = '" & CodigoCuenta & "')"
 FrmEgresos.AdoCordenadas.Refresh
 If FrmEgresos.AdoCordenadas.Recordset.EOF Then
  MsgBox "No Existen Coordenadas, para la Cuenta", vbCritical, "Sistema Contable"
  Exit Sub
 End If
 
 
        X1 = FrmEgresos.AdoCordenadas.Recordset("X1")
        Y1 = FrmEgresos.AdoCordenadas.Recordset("Y1")
        X2 = FrmEgresos.AdoCordenadas.Recordset("X2")
        Y2 = FrmEgresos.AdoCordenadas.Recordset("Y2")
        X3 = FrmEgresos.AdoCordenadas.Recordset("X3")
        Y3 = FrmEgresos.AdoCordenadas.Recordset("Y3")
        X4 = FrmEgresos.AdoCordenadas.Recordset("X4")
        Y4 = FrmEgresos.AdoCordenadas.Recordset("Y4")
        X5 = FrmEgresos.AdoCordenadas.Recordset("X5")
        Y5 = FrmEgresos.AdoCordenadas.Recordset("Y5")
        X6 = FrmEgresos.AdoCordenadas.Recordset("X6")
        Y6 = FrmEgresos.AdoCordenadas.Recordset("Y6")
        X7 = FrmEgresos.AdoCordenadas.Recordset("X7")
        Y7 = FrmEgresos.AdoCordenadas.Recordset("Y7")
        X8 = FrmEgresos.AdoCordenadas.Recordset("X8")
        Y8 = FrmEgresos.AdoCordenadas.Recordset("Y8")
        X9 = FrmEgresos.AdoCordenadas.Recordset("X9")
        Y9 = FrmEgresos.AdoCordenadas.Recordset("Y9")
        X10 = FrmEgresos.AdoCordenadas.Recordset("X10")
        Y10 = FrmEgresos.AdoCordenadas.Recordset("Y10")
        X11 = FrmEgresos.AdoCordenadas.Recordset("X11")
        Y11 = FrmEgresos.AdoCordenadas.Recordset("Y11")
        X12 = FrmEgresos.AdoCordenadas.Recordset("X12")
        Y12 = FrmEgresos.AdoCordenadas.Recordset("Y12")
        X13 = FrmEgresos.AdoCordenadas.Recordset("X13")
        Y13 = FrmEgresos.AdoCordenadas.Recordset("Y13")
        NLineas = Val(FrmEgresos.AdoCordenadas.Recordset("NLineas"))
        CaracteresLineas = Val(FrmEgresos.AdoCordenadas.Recordset("CaracteresLineas"))
        
         If FrmEgresos.CmbMoneda.Text = "Córdobas" Then
            If FrmEgresos.ChkCheque.Value = 1 Then
              TasaCambio = BuscaTasaCambio(FrmEgresos.TxtFecha.Value)
              Monto = FrmEgresos.TxtMonto.Text
              Monto = Monto / TasaCambio
              FrmEgresos.TxtLetras.Text = sw.ConvertCurrencyToSpanish(Monto, "Dólares")
'              ArepEgresos.LblMonto.Caption = Format(Monto, "##,##0.00")
            Else
              Monto = FrmEgresos.TxtMonto.Text
            End If
         Else
           Monto = FrmEgresos.TxtMonto.Text
         End If
 
'Cargo la Consulta del Cheque
 FrmEgresos.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento,Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito,Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas,Transacciones.NumeroMovimiento , Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas1 & "', 102))AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ")ORDER BY Transacciones.NTransaccion"
 FrmEgresos.DtaConsulta.Refresh
 Printer.FontSize = 9
 'Inicio el Ciclo de Impresion
 i = 1
 Do While Not FrmEgresos.DtaConsulta.Recordset.EOF
   
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
                    Printer.CurrentX = Val(X5) '5
                    Printer.CurrentY = Val(Y5) + (5 * i) '120
                    Printer.FontName = "Times New Roman"
                    Printer.FontSize = 11
                    Printer.FontBold = True
                    Printer.Print Concepto
                   End If
                    
                    Dia = Day(FrmEgresos.DtaConsulta.Recordset("FechaTransaccion"))
                    mes = Month(FrmEgresos.DtaConsulta.Recordset("FechaTransaccion"))
                    Año = Year(FrmEgresos.DtaConsulta.Recordset("FechaTransaccion"))
                    Meses = Month(FrmEgresos.DtaConsulta.Recordset("FechaTransaccion"))
                  
                '    FrmEgresos.DtaConsulta.Recordset.MoveLast
                   If X9 <> 0 Or Y9 <> 0 Then
                    Printer.CurrentX = X9
                    Printer.CurrentY = Y9
                    Printer.FontName = "Times New Roman"
                    Printer.FontSize = 11
                    Printer.FontBold = True
                    Printer.Print FrmEgresos.DtaConsulta.Recordset("NumeroMovimiento")
                   End If
                    
                   If X1 <> 0 Or Y1 <> 0 Then
                    Printer.CurrentX = Val(X1)
                    Printer.CurrentY = Val(Y1) + (5 * i)
                    Printer.FontName = "Times New Roman"
                    Printer.FontSize = 11
                    Printer.FontBold = True
                    Printer.Print FrmEgresos.TxtNombre.Text
                   End If
                   
                   If X4 <> 0 Or Y4 <> 0 Then
                    Printer.CurrentX = Val(X4)
                    Printer.CurrentY = Val(Y4) + (5 * i)
                    Printer.FontName = "Times New Roman"
                    Printer.FontSize = 11
                    Printer.FontBold = True
                    Printer.Print FrmEgresos.TxtLetras.Text
                   End If
                   
                   If X3 <> 0 Or Y3 <> 0 Then
                    Printer.CurrentX = Val(X3)
                    Printer.CurrentY = Val(Y3) + (5 * i)
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
                    Printer.Print Format(FrmEgresos.DtaConsulta.Recordset("FechaTransaccion"), "Long date")
                   
'                    Printer.CurrentX = Val(X2) '20
'                    Printer.CurrentY = Val(Y2) '288
'                    Printer.Print Dia
'
'                    Printer.CurrentX = Val(X2) + 18 '38
'                    Printer.CurrentY = Y2 '288
'                    Printer.Print MesLetras(Meses)
'
'                    Printer.CurrentX = Val(X2) + 58 '58
'                    Printer.CurrentY = Y2 '288
'                    Printer.Print Año
                   End If
                
                
           End If
           
           '//////////////////////////////////////////////////////////////////////////////////////
           '//////////////////////////IMPRIMO LOS DETALLES ////////////////////////////////////
           '//////////////////////////////////////////////////////////////////////////////////
           
           
           If X6 <> 0 Or Y6 <> 0 Then
            Printer.CurrentX = Val(X6) '5
            Printer.CurrentY = Val(Y6) + (5 * i)
            Cadena = FrmEgresos.DtaConsulta.Recordset("CodCuentas")
            If Len(Cadena) > 20 Then
             Cadena = Mid(Cadena, 1, 20)
            End If
            
            Printer.FontName = "Times New Roman"
            Printer.FontSize = 9
            Printer.FontBold = False
            Printer.Print Cadena
           End If
        
        
        
        
          If X10 <> 0 Or Y10 <> 0 Then
            Printer.CurrentX = Val(X10) '25
            Printer.CurrentY = Val(Y10) + (5 * i)
            Cadena = FrmEgresos.DtaConsulta.Recordset("NombreCuenta")
            If Len(Cadena) > 24 Then
             Cadena = Mid(Cadena, 1, 24)
            End If
            
            Printer.FontName = "Times New Roman"
            Printer.FontSize = 9
            Printer.FontBold = False
            Printer.Print Cadena
          End If
        
         
            If X11 <> 0 Or Y11 <> 0 Then
                     CadenaDescripcion = FrmEgresos.DtaConsulta.Recordset("DescripcionMovimiento")
                     Cadena = FrmEgresos.DtaConsulta.Recordset("DescripcionMovimiento")
                     Caracter = 1
                     ContadorLinea = i
                     
                     If Len(Cadena) > CaracteresLineas Then
                              Do While Len(Cadena) >= CaracteresLineas
                                       If Caracter = 1 Then
                                                 Cadena = Mid(Cadena, 1, CaracteresLineas)
                                                 Printer.CurrentX = Val(X11) '25
                                                 Printer.CurrentY = Val(Y11) + (5 * i)
                                                 Printer.FontName = "Times New Roman"
                                                 Printer.FontSize = 9
                                                 Printer.FontBold = False
                                                 Printer.Print Cadena
                                                 Caracter = Caracter + CaracteresLineas
                                                 
                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA SIGUIENTE/////////////////
                                                 Cadena = Mid(CadenaDescripcion, Caracter, CaracteresLineas)
                                                 If Len(Cadena) < CaracteresLineas Then
                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                     ContadorLinea = ContadorLinea + 1
                                                     Printer.CurrentX = Val(X11) '25
                                                     Printer.CurrentY = Val(Y11) + (5 * ContadorLinea)
                                                     Printer.FontName = "Times New Roman"
                                                     Printer.FontSize = 9
                                                     Printer.FontBold = False
                                                     Printer.Print Cadena
                                                     
                                                     Caracter = Caracter + CaracteresLineas
                                                 End If
                                         Else
                                                 ContadorLinea = ContadorLinea + 1
                                                 Cadena = Mid(CadenaDescripcion, Caracter, CaracteresLineas)
                                                 Printer.CurrentX = Val(X11) '25
                                                 Printer.CurrentY = Val(Y11) + (5 * ContadorLinea)
                                                 Printer.FontName = "Times New Roman"
                                                 Printer.FontSize = 9
                                                 Printer.FontBold = False
                                                 Printer.Print Cadena
                                                 
                                                 Caracter = Caracter + CaracteresLineas
                                                 
                                                 '//////////////////VERIFICO SI LO QUE SOBRE ES MAYOR DE LA LINEA/////////////////
                                                 Cadena = Mid(CadenaDescripcion, Caracter, CaracteresLineas)
                                                 If Len(Cadena) < CaracteresLineas Then
                                                  '///////////////////////SI ES MENOR IMPRIMO/////////////////////////
                                                     ContadorLinea = ContadorLinea + 1
                                                     Printer.CurrentX = Val(X11) '25
                                                     Printer.CurrentY = Val(Y11) + (5 * ContadorLinea)
                                                     Printer.FontName = "Times New Roman"
                                                     Printer.FontSize = 9
                                                     Printer.FontBold = False
                                                     Printer.Print Cadena
                                                     
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
                             Printer.Print Cadena
                                           
                     End If
                  

            End If
   
              
        
          
           If X12 <> 0 Or Y12 <> 0 Then
            Printer.CurrentX = Val(X12) '135
             Printer.CurrentY = Val(Y12) + (5 * i)
             Printer.FontName = "Times New Roman"
             Printer.FontSize = 9
             Printer.FontBold = False
            Printer.Print Format(FrmEgresos.DtaConsulta.Recordset("Debito"), "##,##0.00")
           End If
        
        
           If X13 <> 0 Or Y13 <> 0 Then
            Printer.CurrentX = Val(X13) '165
              Printer.CurrentY = Val(Y13) + (5 * i) '165
            Printer.FontName = "Times New Roman"
            Printer.FontSize = 9
            Printer.FontBold = False
            Printer.Print Format(FrmEgresos.DtaConsulta.Recordset("Credito"), "##,##0.00")
           End If
           
           
            If i > 1 Then
              UltimaLinea = UltimaLinea + (5 * i) + DiferenciaY - 4
            End If
           
            i = ContadorLinea
            i = i + 1
            ContadorLinea = i
        '
        
        ' 'Fin del Ciclo


 FrmEgresos.DtaConsulta.Recordset.MoveNext
 Loop
'  i = 4
'     i = 4
'    Printer.CurrentX = 70
'    Printer.CurrentY = 140 + (5 * i)
'    Printer.Print "Total"

  If X7 <> 0 Or Y7 <> 0 Then
    Printer.CurrentX = Val(X7) '135
    Printer.CurrentY = Val(Y7) '288
    Printer.FontName = "Times New Roman"
    Printer.FontSize = 9
    Printer.FontBold = False
    Printer.Print Format(TotalDebito, "##,##0.00")
  End If
  
  If X8 <> 0 Or Y8 <> 0 Then
    Printer.CurrentX = Val(X8) '165
    Printer.CurrentY = Val(Y8) '288
    Printer.FontName = "Times New Roman"
    Printer.FontSize = 9
    Printer.FontBold = False
    Printer.Print Format(TotalCredito, "##,##0.00")
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
'    Printer.Print FrmEgresos.TxtNombre.Text
'    Printer.CurrentX = 4
'    Printer.CurrentY = 215 + (5 * i)
'    Printer.Print FrmEgresos.TxtLetras.Text
'    Printer.CurrentX = 140
'    Printer.CurrentY = 206 + (5 * i)
'    Printer.Print Format(FrmEgresos.TxtMonto.Text, "##,##0.00")

   

 
'termino de imprimir las facturas
Printer.EndDoc


' ArepEgresos.DtaCheque.ConnectionString = ConexionReporte
' ArepEgresos.LblMemo = FrmEgresos.TxtMemo
' ArepEgresos.LblNombre2.Caption = FrmEgresos.TxtNombre.Text
' ArepEgresos.LblChequeNo.Caption = Me.LblConsecutivo.Text
' ArepEgresos.Field15.Visible = False
' ArepEgresos.DtaCheque.Source = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento,Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Clave, Transacciones.TCambio, Transacciones.Debito, Transacciones.Credito,Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.Fuente, Transacciones.FechaTasas,Transacciones.NumeroMovimiento , Periodos.Periodo FROM Periodos INNER JOIN Transacciones ON Periodos.NPeriodo = Transacciones.NPeriodo WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fechas1 & "', 102) AND CONVERT(DATETIME, '" & Fechas1 & "', 102))AND (Transacciones.NumeroMovimiento = " & NumeroTransaccion & ")ORDER BY Transacciones.NTransaccion"
' ArepEgresos.Show 1

End If
Unload Me
End Sub
