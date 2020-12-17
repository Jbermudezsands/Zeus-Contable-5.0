VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepSolicitudCheque 
   Caption         =   "Solicitud de Cheque"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepSolicitudCheque.dsx":0000
End
Attribute VB_Name = "ArepSolicitudCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public TipoMoneda As String

Private Sub ActiveReport_ReportStart()
    On Error GoTo err
If Dir(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo) <> "" Then
        Me.Logo.Picture = LoadPicture(MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo)
    End If
err:
    If err.Number <> 0 Then MsgBox "Hay un problema con la dirección del Logo de la Empresa." & Chr(13) & "Por favor revise el valor de la Direccion Logo en la Configuración del Sistema", vbInformation
    
    
   Me.LblEmpresa = MDIPrimero.AdoConfiguracion.Recordset("NombreEmpresa")
   Me.LblEmpresa1 = MDIPrimero.AdoConfiguracion.Recordset("Direccion")
   Me.LblEmpresa2 = "RUC: " & MDIPrimero.AdoConfiguracion.Recordset("NumeroRuc")
   
    Set ew = New cls_NumEnglishWord
    Set sw = New cls_NumSpanishWord
End Sub

Private Sub PageHeader_Format()
 Dim Monto As Double, Letras As String
 
 Monto = Me.FldMonto.Text
 
'
'            If TipoMoneda = "Dólares" Then
'             Letras = sw.ConvertCurrencyToSpanish(Monto, "Dólares")
'            ElseIf TipoMoneda = "Córdobas" Then
'             Letras = sw.ConvertCurrencyToSpanish(Monto, "Córdobas")
'            End If
            
            
 Me.LblDescripcionMonto.Caption = Letras
 
End Sub
