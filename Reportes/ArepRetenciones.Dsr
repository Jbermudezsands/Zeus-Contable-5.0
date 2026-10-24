VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepRetenciones 
   Caption         =   "Retenciones"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35930
   _ExtentY        =   19500
   SectionData     =   "ArepRetenciones.dsx":0000
End
Attribute VB_Name = "ArepRetenciones"
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

    Me.LblMoneda.Caption = Moneda

    '--- Datos de la empresa
    With FrmReportes.DtaDatosEmpresa.Recordset
        Me.LblEmpresa.Caption = .Fields("NombreEmpresa").Value & ""
        Me.LblEmpresa1.Caption = .Fields("Direccion").Value & ""
        Me.LblEmpresa2.Caption = "RUC: " & .Fields("NumeroRuc").Value & ""
    End With

    Me.LblFechaImpreso.Caption = Format$(Now, "dd/mm/yyyy")

    '--- Rango de fechas (y cuenta si se filtró una sola)
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
           "Número: " & err.Number & vbCrLf & _
           "Descripción: " & err.Description, vbExclamation, "Reporte de Retenciones"
    Resume Next

ErrLogo:
    MsgBox "Hay un problema con la dirección del Logo de la Empresa." & vbCrLf & _
           "Por favor revise el valor de Dirección Logo en la Configuración del Sistema.", _
           vbInformation
    Exit Sub

End Sub

