VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepSubReporteLibroMayor 
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12675
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   22357
   _ExtentY        =   13150
   SectionData     =   "ArepSubReporteLibroMayor.dsx":0000
End
Attribute VB_Name = "ArepSubReporteLibroMayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalSaldoInicial As Double, TotalSaldoFinal As Double

Private Sub Detail_Format()
Dim Tipo As String, Debe1 As Double, Haber1 As Double, Debe3 As Double, Haber3 As Double
Dim SaldoInicial As Double, SaldoFinal As Double

If Me.FldDebe1.Text <> "" Then
 Debe1 = Val(Me.FldDebe1.Text)
Else
 Debe1 = 0
End If

If Me.FldHaber1.Text <> "" Then
 Haber1 = Val(Me.FldHaber1.Text)
Else
 Haber1 = 0
End If

If Me.FldDebe3.Text <> "" Then
  Debe3 = Val(Me.FldDebe3.Text)
Else
  Debe3 = 0
End If

If Me.FldHaber3.Text <> "" Then
 Haber3 = Val(Me.FldHaber3.Text)
Else
 Haber3 = 0
End If




Tipo = Mid(Me.FldKeyGrupo.Text, 1, 1)



  If Tipo = "A" Or Tipo = "G" Or Tipo = "O" Then
     SaldoInicial = Debe1 - Haber1
     SaldoFinal = Debe3 - Haber3
     Me.LblSaldoFinal.Caption = Format(SaldoFinal, "##,##0.00")
     Me.LblSaldoInicial.Caption = Format(SaldoInicial, "##,##0.00")
  Else
     SaldoInicial = Haber1 - Debe1
     SaldoFinal = Haber3 - Debe3
     Me.LblSaldoFinal.Caption = Format(SaldoFinal, "##,##0.00")
     Me.LblSaldoInicial.Caption = Format(SaldoInicial, "##,##0.00")
  End If

    TotalSaldoInicial = SaldoInicial + TotalSaldoInicial
    TotalSaldoFinal = SaldoFinal + TotalSaldoFinal
End Sub

Private Sub GroupFooter1_Format()
  Me.LblTotalSaldoInicial.Caption = Format(TotalSaldoInicial, "##,##0.00")
  Me.LblTotalSaldoFinal.Caption = Format(TotalSaldoFinal, "##,##0.00")
  
  TotalInicial = TotalInicial + TotalSaldoInicial
  TotalFinal = TotalFinal + TotalSaldoFinal
  
 If Me.FldTotalDebito.Text <> "" Then
  Debito = Debito + Me.FldTotalDebito.Text
 End If
 If Me.FldTotalCredito.Text <> "" Then
   Credito = Credito + Me.FldTotalCredito.Text
 End If
End Sub


