Attribute VB_Name = "ModCatalogos"
Option Explicit

Public DicCuentas As Object
Public DicCentrosCosto As Object
Public DicTerceros As Object

Public Sub InicializarCatalogos()

    Set DicCuentas = CreateObject("Scripting.Dictionary")
    Set DicCentrosCosto = CreateObject("Scripting.Dictionary")
    Set DicTerceros = CreateObject("Scripting.Dictionary")

End Sub

Public Sub CargarCatalogoCuentas()

    Dim Cn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    Dim Sql As String

    On Error GoTo ErrCargar

    If DicCuentas Is Nothing Then
        InicializarCatalogos
    Else
        DicCuentas.RemoveAll
    End If

    Set Cn = New ADODB.Connection
    Cn.Open Conexion

    Sql = "SELECT CodCuentas FROM Cuentas ORDER BY CodCuentas"

    Set Rs = Cn.Execute(Sql)

    Do While Not Rs.EOF

        If Not DicCuentas.Exists(Trim(Rs!CodCuentas & "")) Then
            DicCuentas.Add Trim(Rs!CodCuentas & ""), True
        End If

        Rs.MoveNext
    Loop

Salir:

    If Not Rs Is Nothing Then
        If Rs.State = adStateOpen Then Rs.Close
    End If

    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then Cn.Close
    End If

    Set Rs = Nothing
    Set Cn = Nothing

    Exit Sub

ErrCargar:

    MsgBox "Error cargando catálogo de cuentas." & vbCrLf & _
           err.Description, vbCritical

    Resume Salir

End Sub


Public Sub CargarCatalogoCentrosCosto()

    Dim Cn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    Dim Sql As String

    On Error GoTo ErrCargar

    If DicCentrosCosto Is Nothing Then
        InicializarCatalogos
    Else
        DicCentrosCosto.RemoveAll
    End If

    Set Cn = New ADODB.Connection
    Cn.Open Conexion

    Sql = "SELECT CodGrupo FROM GrupoCuentas ORDER BY CodGrupo"

    Set Rs = Cn.Execute(Sql)

    Do While Not Rs.EOF

        If Not DicCentrosCosto.Exists(Trim(Rs!CodGrupo & "")) Then
            DicCentrosCosto.Add Trim(Rs!CodGrupo & ""), True
        End If

        Rs.MoveNext
    Loop

Salir:

    If Not Rs Is Nothing Then
        If Rs.State = adStateOpen Then Rs.Close
    End If

    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then Cn.Close
    End If

    Set Rs = Nothing
    Set Cn = Nothing

    Exit Sub

ErrCargar:

    MsgBox "Error cargando catálogo de Centros de Costo." & vbCrLf & _
           err.Description, vbCritical

    Resume Salir

End Sub


Public Sub CargarCatalogoTerceros()

    Dim Cn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    Dim Sql As String

    On Error GoTo ErrCargar

    If DicTerceros Is Nothing Then
        InicializarCatalogos
    Else
        DicTerceros.RemoveAll
    End If

    Set Cn = New ADODB.Connection
    Cn.Open ConexionFacturacion

    Sql = "SELECT Codigo_Beneficiario FROM Beneficiario WHERE Activo = 1"

    Set Rs = Cn.Execute(Sql)

    Do While Not Rs.EOF

        If Not DicTerceros.Exists(Trim(Rs!Codigo_Beneficiario & "")) Then
            DicTerceros.Add Trim(Rs!Codigo_Beneficiario & ""), True
        End If

        Rs.MoveNext
    Loop

Salir:

    If Not Rs Is Nothing Then
        If Rs.State = adStateOpen Then Rs.Close
    End If

    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then Cn.Close
    End If

    Set Rs = Nothing
    Set Cn = Nothing

    Exit Sub

ErrCargar:

    MsgBox "Error cargando catálogo de terceros." & vbCrLf & _
           err.Description, vbCritical

    Resume Salir

End Sub
