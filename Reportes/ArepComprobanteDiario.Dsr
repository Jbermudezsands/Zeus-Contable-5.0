VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepComprobanteDiario 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepComprobanteDiario.dsx":0000
End
Attribute VB_Name = "ArepComprobanteDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoCuenta As String

Private Sub ActiveReport_Activate()
QuienReporte = Me.Name
End Sub

Private Sub ActiveReport_ReportStart()
QuienReporte = Me.Name

'    Me.LblEmpresa = MDIPrimero.AdoConfiguracion.Recordset("NombreEmpresa")
'    ArepTransacciones.LblEmpresa1 = MDIPrimero.AdoConfiguracion.Recordset("Direccion")
'    ArepTransacciones.LblEmpresa2 = "RUC: " & MDIPrimero.AdoConfiguracion.Recordset("NumeroRuc")
''    ArepTransacciones.Logo.Picture = LoadPicture(RutaLogo)
'    ArepTransacciones.LblFecha.Caption = Format(Now, "long")
    
             Me.LblEmpresa = FrmReportes.DtaDatosEmpresa.Recordset("NombreEmpresa")
             Me.LblEmpresa1 = FrmReportes.DtaDatosEmpresa.Recordset("Direccion")
             Me.LblEmpresa2 = "RUC: " & FrmReportes.DtaDatosEmpresa.Recordset("NumeroRuc")
             Me.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
             Me.LblMoneda.Caption = FrmReportes.CmbMoneda.Text
             
             
'             Me.LblMoneda.Caption = "Expresado en " & Moneda
             
'             Me.LblCodigo.Caption = "Mis Code:" & FrmReportes.DBCodigo.Text
'             Me.LblRango.Caption = "Filtrado Desde: " & FrmReportes.CodDesde & " Hasta " & FrmReportes.CodHasta

  
If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    End If
err:
    If err.Number <> 0 Then MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation


End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)
    If Not EOF Then
    'Gets the current records SupplierID
        If Not IsNull(Me.DataControl1.Recordset.Fields("CodCuentas")) Then
          CodigoCuenta = Me.DataControl1.Recordset.Fields("CodCuentas")
        Else
          CodigoCuenta = ""
        End If
    End If
End Sub

Private Sub Detail_Format()
If FrmReportes.ChkExportar.Value = 0 Then
  Me.FldCuenta.Hyperlink = Me.FldCuenta.Text
  'Me.FldCuenta.Style = "color: rgb(0,0,0); font-weight: Arial Narrow; font-size: 9pt"
  Me.FldCuenta.ForeColor = &HC00000
End If
End Sub

Private Sub GroupHeader1_Format()
If Me.Field16.Text <> "" Then
 Me.LblNumeroComprobante.Caption = Year(Format(Me.Field16.Text, "dd/mm/yyyy")) & Month(Format(Me.Field16.Text, "dd/mm/yyyy")) & Day(Format(Me.Field16.Text, "dd/mm/yyyy"))
End If
End Sub

