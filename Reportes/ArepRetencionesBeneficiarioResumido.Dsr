VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepRetencionesBeneficiarioResumido 
   Caption         =   "Retenciones por Beneficiario Resumido"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35930
   _ExtentY        =   19500
   SectionData     =   "ArepRetencionesBeneficiarioResumido.dsx":0000
End
Attribute VB_Name = "ArepRetencionesBeneficiarioResumido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FechaDesde As Date
Public FechaHasta As Date
Public CuentaDesde As String
Public CuentaHasta As String



Private Sub ActiveReport_ReportEnd()

    On Error Resume Next

    If Not DataControl1.Recordset Is Nothing Then
        If DataControl1.Recordset.State <> 0 Then DataControl1.Recordset.Close
    End If

End Sub

Private Sub ActiveReport_ReportStart()

    Dim RutaImg As String

    On Error GoTo ErrDatos

    QuienReporte = Me.Name

    '--- Layout base copiado de ArepRetenciones: resumen, un renglon por beneficiario
    Me.lblTitulo.Caption = "RETENCIONES X BENEFICIARIO RESUMIDO"

    '    Encabezado de columnas
    Me.Label18.Visible = False
    Me.Label4.Visible = False
    Me.Label6.Caption = "Cedula"
    Me.Label27.Caption = "Tasa Ret"
    Me.Label28.Visible = False
    Me.Label5.Caption = "Codigo Benef."
    Me.Label26.Caption = "Nombre Beneficiario"
    Me.Label22.Caption = "Total Retenido"

    '    Detalle
    Me.Field19.DataField = "Fecha"
    Me.Field23.DataField = "Documento"
    Me.Field24.DataField = "Cod_Tercero"
    Me.Field26.DataField = "NombreBeneficiario"
    Me.Field29.DataField = "Retenido"
    Me.Field19.Visible = False
    Me.Field23.Visible = False
    Me.Field25.DataField = "Cedula"
    Me.Field27.DataField = "TasaRet"
    Me.Field28.Visible = False

    '    Reacomodar columnas: Codigo, Cedula, Nombre (espacio libre), Tasa, Total
    Me.Field24.Left = Me.Field19.Left
    Me.Field25.Left = Me.Field24.Left + Me.Field24.Width
    Me.Field25.Width = Me.Field23.Width * 2.2
    Me.Field27.Width = Me.Field24.Width
    Me.Field27.Left = Me.Field29.Left - Me.Field27.Width
    Me.Field26.Left = Me.Field25.Left + Me.Field25.Width
    Me.Field26.Width = Me.Field27.Left - Me.Field26.Left
    Me.Label5.Left = Me.Field24.Left
    Me.Label6.Left = Me.Field25.Left
    Me.Label6.Width = Me.Field25.Width
    Me.Label26.Left = Me.Field26.Left
    Me.Label26.Width = Me.Field26.Width
    Me.Label27.Left = Me.Field27.Left
    Me.Label27.Width = Me.Field27.Width
    '    Sin agrupar: el grupo (constante) solo sirve para el total general del pie
    Me.Sections("GroupHeader1").DataField = "Grupo"
    Me.Sections("GroupHeader1").Visible = False
    Me.Field16.DataField = "Cod_Tercero"
    Me.Field17.DataField = "NombreBeneficiario"
    Me.Field18.DataField = "DescRetencion"
    Me.Field30.Visible = False
    Me.Field32.Visible = False

    Me.LblMoneda.Caption = Moneda

    '--- Datos de la empresa
    With FrmReportes.DtaDatosEmpresa.Recordset
        Me.LblEmpresa.Caption = .Fields("NombreEmpresa").Value & ""
        Me.LblEmpresa1.Caption = .Fields("Direccion").Value & ""
        Me.LblEmpresa2.Caption = "RUC: " & .Fields("NumeroRuc").Value & ""
    End With

    Me.LblFechaImpreso.Caption = Format$(Now, "dd/mm/yyyy")

    '--- Rango de fechas (y cuentas si se filtraron)
    Me.LblRango.Caption = "Del " & Format$(FechaDesde, "dd/mm/yyyy") & _
                          " al " & Format$(FechaHasta, "dd/mm/yyyy")

    If Len(CuentaDesde) > 0 Or Len(CuentaHasta) > 0 Then
        Me.LblRango.Caption = Me.LblRango.Caption & _
            "   Cuentas: " & IIf(Len(CuentaDesde) > 0, CuentaDesde, "Inicio") & _
            " - " & IIf(Len(CuentaHasta) > 0, CuentaHasta, "Final")
    End If

    '--- Logo (error independiente, no detiene el reporte)
    On Error GoTo ErrLogo

    RutaImg = MDIPrimero.AdoConfiguracion.Recordset!DireccionLogo & ""

    If Len(RutaImg) > 0 Then
        If Dir(RutaImg) <> "" Then
            Me.Logo.Picture = LoadPicture(RutaImg)
        End If
    End If

    Exit Sub

ErrDatos:
    MsgBox "Error al cargar los datos del encabezado del reporte." & vbCrLf & _
           "Numero: " & err.Number & vbCrLf & _
           "Descripcion: " & err.Description, vbExclamation, "Retenciones por Beneficiario"
    Resume Next

ErrLogo:
    MsgBox "Hay un problema con la direccion del Logo de la Empresa." & vbCrLf & _
           "Por favor revise el valor de Direccion Logo en la Configuracion del Sistema.", _
           vbInformation
    Exit Sub

End Sub
