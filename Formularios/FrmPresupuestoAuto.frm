VERSION 5.00
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Begin VB.Form FrmPresupuestoAuto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculo Automatico"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   Icon            =   "FrmPresupuestoAuto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data DtaConsulta 
      Caption         =   "DtaConsulta"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Width           =   2655
   End
   Begin SmartButtonProject.SmartButton CmdCancelar 
      Height          =   735
      Left            =   4680
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      Caption         =   "Cancelar"
      Picture         =   "FrmPresupuestoAuto.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureLayout   =   7
   End
   Begin SmartButtonProject.SmartButton CmdAceptar 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      Caption         =   "Aceptar"
      Picture         =   "FrmPresupuestoAuto.frx":0BE4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureLayout   =   7
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.Frame Frame2 
         Caption         =   "Periodos"
         Height          =   615
         Left            =   2640
         TabIndex        =   8
         Top             =   120
         Width           =   2895
         Begin VB.OptionButton Option6 
            Caption         =   "2003"
            Height          =   255
            Left            =   1920
            TabIndex        =   11
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option5 
            Caption         =   "2002"
            Height          =   255
            Left            =   960
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option4 
            Caption         =   "2001"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.TextBox TxtPorciento 
         Height          =   285
         Left            =   3120
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Saldo Año Anterior"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Presupuesto Año Anterior"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Importe Anual"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Importe Anual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   840
         Width           =   2895
      End
   End
   Begin vbskfree.Skinner Skinner1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
End
Attribute VB_Name = "FrmPresupuestoAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmdAceptar_Click()
Dim NumeroP As Integer, I As Integer
Dim Monto As Double, MontoPre As Double
Dim Diferencia As Double, CodCuenta As String
Dim Porciento As Double, Debito As Double, Credito As Double, Total1 As Double

If Not IsNumeric(Me.TxtPorciento.Text) Then
 MsgBox "El Campo Digitado no es Numerico", vbCritical, "Sistema contable"
 Exit Sub
Else
 Monto = Me.TxtPorciento.Text
End If

If Me.TxtPorciento.Text = "" Then
  MsgBox "No puede dejar en Blanco el Campo " & Me.Label1, vbInformation, "sistema Contable"
  Exit Sub
End If

   If Me.Option4 Then
     NumeroP = 1
   ElseIf Me.Option5 Then
     NumeroP = 2
   ElseIf Me.Option6 Then
     NumeroP = 3
   End If

 If Me.Option1 Then
  MontoPre = 0
     FrmPresupuesto.Text1.Text = Format(Monto / 12, "##,##0.00")
     FrmPresupuesto.Text2.Text = Format(Monto / 12, "##,##0.00")
     FrmPresupuesto.Text3.Text = Format(Monto / 12, "##,##0.00")
     FrmPresupuesto.Text4.Text = Format(Monto / 12, "##,##0.00")
     FrmPresupuesto.Text5.Text = Format(Monto / 12, "##,##0.00")
     FrmPresupuesto.Text6.Text = Format(Monto / 12, "##,##0.00")
     FrmPresupuesto.Text7.Text = Format(Monto / 12, "##,##0.00")
     FrmPresupuesto.Text8.Text = Format(Monto / 12, "##,##0.00")
     FrmPresupuesto.Text9.Text = Format(Monto / 12, "##,##0.00")
     FrmPresupuesto.Text10.Text = Format(Monto / 12, "##,##0.00")
     FrmPresupuesto.Text11.Text = Format(Monto / 12, "##,##0.00")
     FrmPresupuesto.Text12.Text = Format(Monto / 12, "##,##0.00")
     MontoPre = FrmPresupuesto.TxtTotal1.Text

       Diferencia = (Monto - MontoPre)
       FrmPresupuesto.Text12.Text = FrmPresupuesto.Text12.Text + Diferencia
   
 ElseIf Me.Option2 Then
 
 MontoPre = 0
 Porciento = 1 + (Monto / 100)
 If NumeroP = 1 Then
   MsgBox "No se puede Presupuestar para este Año, seleccione el siguiente", vbCritical, "Sistema contable"
   Exit Sub
 End If
 
 FrmPresupuesto.DtaPeriodos.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos Where (((Periodos.NumeroTabla) = " & NumeroP - 1 & "))"
 FrmPresupuesto.DtaPeriodos.Refresh
 Do While Not FrmPresupuesto.DtaPeriodos.Recordset.EOF
  NumeroPeriodo = FrmPresupuesto.DtaPeriodos.Recordset.NPeriodo
  
  Periodo = FrmPresupuesto.DtaPeriodos.Recordset.Periodo
  CodigoCuenta = FrmPresupuesto.DBCliente.Text
  Me.DtaConsulta.RecordSource = "SELECT Presupuesto.NPeriodo, Presupuesto.CodCuenta, Presupuesto.MontoPresupuestado, Presupuesto.SaldoReal From Presupuesto Where (((Presupuesto.NPeriodo) = " & NumeroPeriodo & " ) And ((Presupuesto.CodCuenta) = '" & CodigoCuenta & "'))"
  Me.DtaConsulta.Refresh
  If Not Me.DtaConsulta.Recordset.EOF Then
   Saldo = Me.DtaConsulta.Recordset.MontoPresupuestado
   MontoPre = Saldo + MontoPre

  Else
   Saldo = 0#
 End If
  
  If NumeroP = 2 Then
    Select Case Periodo
      Case 1: FrmPresupuesto.Text13 = Format(Saldo * Porciento, "##,##0.00")
      Case 2: FrmPresupuesto.Text14 = Format(Saldo * Porciento, "##,##0.00")
      Case 3: FrmPresupuesto.Text15 = Format(Saldo * Porciento, "##,##0.00")
      Case 4: FrmPresupuesto.Text16 = Format(Saldo * Porciento, "##,##0.00")
      Case 5: FrmPresupuesto.Text17 = Format(Saldo * Porciento, "##,##0.00")
      Case 6: FrmPresupuesto.Text18 = Format(Saldo * Porciento, "##,##0.00")
      Case 7: FrmPresupuesto.Text19 = Format(Saldo * Porciento, "##,##0.00")
      Case 8: FrmPresupuesto.Text20 = Format(Saldo * Porciento, "##,##0.00")
      Case 9: FrmPresupuesto.Text21 = Format(Saldo * Porciento, "##,##0.00")
      Case 10: FrmPresupuesto.Text22 = Format(Saldo * Porciento, "##,##0.00")
      Case 11: FrmPresupuesto.Text23 = Format(Saldo * Porciento, "##,##0.00")
      Case 12: FrmPresupuesto.Text24 = Format(Saldo * Porciento, "##,##0.00")
               Monto = FrmPresupuesto.TxtTotal2.Text
               Diferencia = (MontoPre * Porciento - Monto)
               Saldo = Saldo + Diferencia
               FrmPresupuesto.Text24 = Format(Saldo * Porciento, "##,##0.00")

    End Select
   ElseIf NumeroP = 3 Then
    Select Case Periodo
      Case 1: FrmPresupuesto.Text25 = Format(Saldo * Porciento, "##,##0.00")
      Case 2: FrmPresupuesto.Text26 = Format(Saldo * Porciento, "##,##0.00")
      Case 3: FrmPresupuesto.Text27 = Format(Saldo * Porciento, "##,##0.00")
      Case 4: FrmPresupuesto.Text28 = Format(Saldo * Porciento, "##,##0.00")
      Case 5: FrmPresupuesto.Text29 = Format(Saldo * Porciento, "##,##0.00")
      Case 6: FrmPresupuesto.Text30 = Format(Saldo * Porciento, "##,##0.00")
      Case 7: FrmPresupuesto.Text31 = Format(Saldo * Porciento, "##,##0.00")
      Case 8: FrmPresupuesto.Text32 = Format(Saldo * Porciento, "##,##0.00")
      Case 9: FrmPresupuesto.Text33 = Format(Saldo * Porciento, "##,##0.00")
      Case 10: FrmPresupuesto.Text34 = Format(Saldo * Porciento, "##,##0.00")
      Case 11: FrmPresupuesto.Text35 = Format(Saldo * Porciento, "##,##0.00")
      Case 12: FrmPresupuesto.Text36.Text = Format(Saldo * Porciento, "##,##0.00")
               Monto = FrmPresupuesto.TxtTotal3.Text
               Diferencia = (MontoPre * Porciento - Monto)
               Saldo = Saldo + Diferencia
               FrmPresupuesto.Text36 = Format(Saldo * Porciento, "##,##0.00")
    End Select
   
   End If
    
  FrmPresupuesto.DtaPeriodos.Recordset.MoveNext
 Loop
 
 


 ElseIf Me.Option3 Then
 
  MontoPre = 0
 Porciento = 1 + (Monto / 100)
 If NumeroP = 1 Then
   MsgBox "No se puede Presupuestar para este Año, seleccione el siguiente", vbCritical, "Sistema contable"
   Exit Sub
 End If
 
 FrmPresupuesto.DtaPeriodos.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos Where (((Periodos.NumeroTabla) = " & NumeroP - 1 & "))"
 FrmPresupuesto.DtaPeriodos.Refresh
 Do While Not FrmPresupuesto.DtaPeriodos.Recordset.EOF
  NumeroPeriodo = FrmPresupuesto.DtaPeriodos.Recordset.NPeriodo
  
  Periodo = FrmPresupuesto.DtaPeriodos.Recordset.Periodo
  CodigoCuenta = FrmPresupuesto.DBCliente.Text
  Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Sum([Transacciones]![Debito]*[Transacciones]![TCambio]) AS MDebito, Sum([Transacciones]![TCambio]*[Transacciones]![Credito]) AS MCredito, Transacciones.NPeriodo From Transacciones GROUP BY Transacciones.CodCuentas, Transacciones.NPeriodo HAVING (((Transacciones.CodCuentas)='" & CodigoCuenta & "') AND ((Transacciones.NPeriodo)=" & NumeroPeriodo & "))"
  Me.DtaConsulta.Refresh

  If Not Me.DtaConsulta.Recordset.EOF Then
    If Not IsNull(Me.DtaConsulta.Recordset.MDebito) Then
    Debito = Me.DtaConsulta.Recordset.MDebito
    Else
     Debito = 0
   End If
  If Not IsNull(Me.DtaConsulta.Recordset.MCredito) Then
   Credito = Me.DtaConsulta.Recordset.MCredito
  Else
   Credito = 0
  End If
   Saldo = (Debito - Credito)
   Total1 = Total1 + Saldo
  Else
   Saldo = 0#
 End If
  
  If NumeroP = 2 Then
    Select Case Periodo
      Case 1: FrmPresupuesto.Text13 = Format(Saldo * Porciento, "##,##0.00")
      Case 2: FrmPresupuesto.Text14 = Format(Saldo * Porciento, "##,##0.00")
      Case 3: FrmPresupuesto.Text15 = Format(Saldo * Porciento, "##,##0.00")
      Case 4: FrmPresupuesto.Text16 = Format(Saldo * Porciento, "##,##0.00")
      Case 5: FrmPresupuesto.Text17 = Format(Saldo * Porciento, "##,##0.00")
      Case 6: FrmPresupuesto.Text18 = Format(Saldo * Porciento, "##,##0.00")
      Case 7: FrmPresupuesto.Text19 = Format(Saldo * Porciento, "##,##0.00")
      Case 8: FrmPresupuesto.Text20 = Format(Saldo * Porciento, "##,##0.00")
      Case 9: FrmPresupuesto.Text21 = Format(Saldo * Porciento, "##,##0.00")
      Case 10: FrmPresupuesto.Text22 = Format(Saldo * Porciento, "##,##0.00")
      Case 11: FrmPresupuesto.Text23 = Format(Saldo * Porciento, "##,##0.00")
      Case 12: FrmPresupuesto.Text24 = Format(Saldo * Porciento, "##,##0.00")
               Monto = FrmPresupuesto.TxtTotal2.Text
               Diferencia = (Total1 * Porciento - Monto)
               Saldo = Saldo + Diferencia
               FrmPresupuesto.Text24 = Format(Saldo * Porciento, "##,##0.00")

    End Select
   ElseIf NumeroP = 3 Then
    Select Case Periodo
      Case 1: FrmPresupuesto.Text25 = Format(Saldo * Porciento, "##,##0.00")
      Case 2: FrmPresupuesto.Text26 = Format(Saldo * Porciento, "##,##0.00")
      Case 3: FrmPresupuesto.Text27 = Format(Saldo * Porciento, "##,##0.00")
      Case 4: FrmPresupuesto.Text28 = Format(Saldo * Porciento, "##,##0.00")
      Case 5: FrmPresupuesto.Text29 = Format(Saldo * Porciento, "##,##0.00")
      Case 6: FrmPresupuesto.Text30 = Format(Saldo * Porciento, "##,##0.00")
      Case 7: FrmPresupuesto.Text31 = Format(Saldo * Porciento, "##,##0.00")
      Case 8: FrmPresupuesto.Text32 = Format(Saldo * Porciento, "##,##0.00")
      Case 9: FrmPresupuesto.Text33 = Format(Saldo * Porciento, "##,##0.00")
      Case 10: FrmPresupuesto.Text34 = Format(Saldo * Porciento, "##,##0.00")
      Case 11: FrmPresupuesto.Text35 = Format(Saldo * Porciento, "##,##0.00")
      Case 12: FrmPresupuesto.Text36.Text = Format(Saldo * Porciento, "##,##0.00")
               Monto = FrmPresupuesto.TxtTotal3.Text
               Diferencia = (Total1 * Porciento - Monto)
               Saldo = Saldo + Diferencia
               FrmPresupuesto.Text36 = Format(Saldo * Porciento, "##,##0.00")
    End Select
   
   End If
    
  FrmPresupuesto.DtaPeriodos.Recordset.MoveNext
 Loop
 
 End If
 
 Unload Me

End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim AÑO1 As String, AÑO2 As String, AÑO3 As String

With Me.DtaConsulta
    .DatabaseName = Ruta
    .Connect = Conexion
End With

      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
      Me.DtaConsulta.Refresh
      Do While Not DtaConsulta.Recordset.EOF
       If AÑO1 = "" Then
        AÑO1 = Year(DtaConsulta.Recordset.FechaPeriodo)
        Me.Option4.Caption = AÑO1
       ElseIf AÑO2 = "" Then
        AÑO2 = Year(DtaConsulta.Recordset.FechaPeriodo)
        Me.Option5.Caption = AÑO2
       Else
         AÑO3 = Year(DtaConsulta.Recordset.FechaPeriodo)
         Me.Option6.Caption = AÑO3
       End If
        
        Me.DtaConsulta.Recordset.MoveNext
      Loop
      

      
End Sub

Private Sub Option1_Click()
Me.Label1.Caption = "Importe Anual"
End Sub

Private Sub Option2_Click()
Me.Label1.Caption = "Porciento Incremento"
End Sub

Private Sub Option3_Click()
Me.Label1.Caption = "Porciento Incremento"
End Sub
