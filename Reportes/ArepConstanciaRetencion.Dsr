VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepConstanciaRetencion 
   Caption         =   "Constancia de Retencion"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepConstanciaRetencion.dsx":0000
End
Attribute VB_Name = "ArepConstanciaRetencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
   On Error GoTo err
    
    QuienReporte = Me.Name
    
             Me.LblEmpresa = MDIPrimero.AdoConfiguracion.Recordset("NombreEmpresa")
             Me.LblEmpresa1 = MDIPrimero.AdoConfiguracion.Recordset("Direccion")
             Me.LblEmpresa2 = "RUC: " & MDIPrimero.AdoConfiguracion.Recordset("NumeroRuc")
             If Dir(RutaLogo) <> "" Then
                   Me.Logo.Picture = LoadPicture(RutaLogo)
             End If

    
    
    
If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    End If
err:
    If err.Number <> 0 Then MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation
End Sub

