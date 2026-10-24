VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepRetencionesBeneficiario 
   Caption         =   "Retenciones por Beneficiario"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35930
   _ExtentY        =   19500
   SectionData     =   "ArepRetencionesBeneficiario.dsx":0000
End
Attribute VB_Name = "ArepRetencionesBeneficiario"
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

    '--- Layout base copiado de ArepRetenciones: se reasignan titulos y campos
    Me.lblTitulo.Caption = "RETENCIONES X BENEFICIARIO DETALLE"

    '    Encabezado de columnas
    Me.Label4.Caption = "Documento"
    Me.Label5.Caption = "Codigo Benef."
    Me.Label26.Caption = "Nombre Beneficiario"
    Me.Label22.Caption = "Monto Retenido"
    Me.Label6.Visible = False
    Me.Label27.Visible = False
    Me.Label28.Visible = False

    '    Detalle
    Me.Field19.DataField = "Fecha"
    Me.Field19.OutputFormat = "dd/mm/yyyy"
    Me.Field23.DataField = "Documento"
    Me.Field24.DataField = "Cod_Tercero"
    Me.Field26.DataField = "NombreBeneficiario"
    Me.Field29.DataField = "Retenido"
    Me.Field25.Visible = False
    Me.Field27.Visible = False
    Me.Field28.Visible = False

    '    Grupo por beneficiario (antes por cuenta)
    Me.Sections("GroupHeader1").DataField = "Cod_Tercero"
    Me.Label29.Caption = "Codigo"
    Me.Label30.Caption = "Nombre"
    Me.Field16.DataField = "Cod_Tercero"
    Me.Field17.DataField = "NombreBeneficiario"
    Me.Field18.Visible = False

    '    Pie de grupo: solo total retenido por beneficiario
    Me.Field30.Visible = False
    Me.Field32.Visible = False

    '    Reacomodar columnas en una sola pasada: Fecha, Documento (mas ancho),
    '    Codigo y Nombre (ocupa el espacio libre hasta el Monto)
    Me.Field23.Left = Me.Field19.Left + Me.Field19.Width
    Me.Field23.Width = Me.Field23.Width * 1.6
    Me.Field24.Left = Me.Field23.Left + Me.Field23.Width
    Me.Field26.Left = Me.Field24.Left + Me.Field24.Width
    Me.Field26.Width = Me.Field29.Left - Me.Field26.Left

    Me.Label18.Left = Me.Field19.Left
    Me.Label4.Left = Me.Field23.Left
    Me.Label4.Width = Me.Field23.Width
    Me.Label5.Left = Me.Field24.Left
    Me.Label26.Left = Me.Field26.Left
    Me.Label26.Width = Me.Field26.Width

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
