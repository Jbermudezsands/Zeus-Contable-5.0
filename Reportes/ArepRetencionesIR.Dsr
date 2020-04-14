VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepRetencionesIR 
   Caption         =   "Reporte Retenciones de la Fuente IR"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   33655
   _ExtentY        =   20267
   SectionData     =   "ArepRetencionesIR.dsx":0000
End
Attribute VB_Name = "ArepRetencionesIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalPago As Double

Private Sub Detail_Format()
 Dim CodCuenta As String, NumeroFactura As String
 
 If Not Me.FldRetencion.Text = "" Then
 TotalPago = TotalPago + CDbl(Me.FldRetencion.Text)
 End If
 CodCuenta = Me.FldCuenta.Text
 NumeroFactura = Me.FldFactura.Text
 FrmReportes.DtaConsulta.RecordSource = "SELECT *, Transacciones.TCambio * (Transacciones.Debito + Transacciones.Credito) AS MontoFactura " & _
                                        "FROM Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas  " & _
                                        "WHERE  (Cuentas.CausaIva <> 1) AND (Cuentas.CausaRetencion <> 1) AND (Transacciones.FacturaNo = '" & NumeroFactura & "')"
 FrmReportes.DtaConsulta.Refresh
 If Not FrmReportes.DtaConsulta.Recordset.EOF Then
  Me.LblMontoFactura.Caption = Format(FrmReportes.DtaConsulta.Recordset("MontoFactura"), "##,##0.00")
 
 End If
 
 FrmReportes.DtaConsulta2.RecordSource = "SELECT * From Cuentas WHERE (CodCuentas = '" & CodCuenta & "')"
 FrmReportes.DtaConsulta2.Refresh
 If Not FrmReportes.DtaConsulta2.Recordset.EOF Then
   If Not IsNull(FrmReportes.DtaConsulta2.Recordset("RUC")) Then
    Me.LblRUCS.Caption = FrmReportes.DtaConsulta2.Recordset("RUC")
   End If
   
   If Not IsNull(FrmReportes.DtaConsulta2.Recordset("Cedula")) Then
    Me.LblCedula.Caption = FrmReportes.DtaConsulta2.Recordset("Cedula")
   End If
   
   If Not IsNull(FrmReportes.DtaConsulta2.Recordset("Apellido1")) Then
    Me.LblApellido1.Caption = FrmReportes.DtaConsulta2.Recordset("Apellido1")
   End If
   
  If Not IsNull(FrmReportes.DtaConsulta2.Recordset("Apellido2")) Then
    Me.LblApellido2.Caption = FrmReportes.DtaConsulta2.Recordset("Apellido2")
   End If
   
   If Not IsNull(FrmReportes.DtaConsulta2.Recordset("Nombre1")) Then
    Me.LblNombres.Caption = FrmReportes.DtaConsulta2.Recordset("Nombre1")
    If Not IsNull(FrmReportes.DtaConsulta2.Recordset("Nombre2")) Then
     Me.LblNombres.Caption = FrmReportes.DtaConsulta2.Recordset("Nombre1") + " " + FrmReportes.DtaConsulta2.Recordset("Nombre2")
    End If
   End If
   
   If Not IsNull(FrmReportes.DtaConsulta2.Recordset("DescripcionCuentas")) Then
    Me.LblNombreComercial.Caption = FrmReportes.DtaConsulta2.Recordset("DescripcionCuentas")
   End If
 
 End If

End Sub

Private Sub ReportFooter_Format()
 Me.LblTotalPago.Caption = Format(TotalPago, "##,##0.00")
End Sub
