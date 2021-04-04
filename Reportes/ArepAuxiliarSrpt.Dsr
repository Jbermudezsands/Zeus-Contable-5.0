VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepAuxiliarSrpt 
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   19080
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   33655
   _ExtentY        =   19368
   SectionData     =   "ArepAuxiliarSrpt.dsx":0000
End
Attribute VB_Name = "ArepAuxiliarSrpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TransaccionNo As String, FechaTransaccion As Date, Saldo As Double
Private Sub ActiveReport_FetchData(EOF As Boolean)
    If Not EOF Then
    'Gets the current records SupplierID
      TransaccionNo = Me.DataControl1.Recordset.Fields("NumeroMovimiento")
      FechaTransaccion = Me.DataControl1.Recordset.Fields("FechaTransaccion")
    End If
End Sub

Private Sub Detail_Format()
Dim Mov1 As Double, Mov2 As Double

  

 If Me.Field26.Text = "0.00" Or Me.Field26.Text = "" Then
   Mov1 = 0
 Else
   Mov1 = Me.Field26.Text
 End If
 
  If Me.Field27.Text = "0.00" Or Me.Field27.Text = "" Then
   Mov2 = 0
 Else
   Mov2 = Me.Field27.Text
 End If
 
      TipoCuenta = Me.FldTipoCuenta.Text
    

If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
 Saldo = Mov1 - Mov2 + Saldo
    Me.FldSaldo.Text = Format(SaldoIni + Saldo, "###,###,###,##0.#0")
    SaldoFinalAuxiliar = Format(SaldoIni + Saldo, "###,###,###,##0.#0")
Else
 Saldo = Mov2 - Mov1 + Saldo
    Me.FldSaldo.Text = Format(SaldoIni + Saldo, "###,###,###,##0.#0")
    SaldoFinalAuxiliar = Format(SaldoIni + Saldo, "###,###,###,##0.#0")
'    Me.LblFinal.Caption = Format(SaldoIni - SaldoFin, "##,##0.00")
End If

'pone al tipo sólo D o C
Me.Field22.Text = Mid(Me.Field22, 1, 1)
If FrmReportes.ChkExportar.Value = 0 Then
  Me.Field19.Hyperlink = Format(FechaTransaccion, "dd/mm/yyyy") & ";" & TransaccionNo
End If


End Sub
