Attribute VB_Name = "Funciones2"
Option Explicit

Public AuxiliarDetallado As Boolean

'==========================================================
' SQL del auxiliar de cuentas (detallado o totalizado)
'   MonedaReporte : "Córdobas" o "Dólares"
'   Rangos vacíos = sin filtro
'   CodBeneficiario vacío o "TODOS" = sin filtro
' Columnas extra: KeyGrupo, DescripcionGrupo, DebitoAnterior, CreditoAnterior
'==========================================================
Public Function SqlAuxiliarCuentas( _
        ByVal FechaIni As Date, _
        ByVal FechaFin As Date, _
        ByVal MonedaReporte As String, _
        Optional ByVal CuentaDesde As String = "", _
        Optional ByVal CuentaHasta As String = "", _
        Optional ByVal CodBeneficiario As String = "", _
        Optional ByVal Detallado As Boolean = True, _
        Optional ByVal AgruparPorBeneficiario As Boolean = False, _
        Optional ByVal GrupoDesde As String = "", _
        Optional ByVal GrupoHasta As String = "") As String

    Dim Sql As String, F1 As String, F2 As String
    Dim MonedaSql As String, AjusteExcluir As String, OperadorTasa As String
    Dim ExprDebito As String, ExprCredito As String
    Dim FiltroCuentas As String, FiltroBenef As String, FiltroGrupos As String
    Dim JoinIndice As String, JoinTasas As String

    '--- Moneda
    If MonedaReporte = "Córdobas" Then
        MonedaSql = "Córdobas": AjusteExcluir = "Dólares": OperadorTasa = " * "
    Else
        MonedaSql = "Dólares": AjusteExcluir = "Córdobas": OperadorTasa = " / "
    End If

    F1 = Format(FechaIni, "yyyy-mm-dd")
    F2 = Format(FechaFin, "yyyy-mm-dd")

    '--- Filtros opcionales
    If Trim$(CuentaDesde) <> "" Then FiltroCuentas = FiltroCuentas & _
        "AND Transacciones.CodCuentas >= '" & Replace(Trim$(CuentaDesde), "'", "''") & "' "
    If Trim$(CuentaHasta) <> "" Then FiltroCuentas = FiltroCuentas & _
        "AND Transacciones.CodCuentas <= '" & Replace(Trim$(CuentaHasta), "'", "''") & "' "

    CodBeneficiario = Trim$(CodBeneficiario)
    If CodBeneficiario <> "" And UCase$(CodBeneficiario) <> "TODOS" Then
        FiltroBenef = "AND Transacciones.Cod_Tercero = '" & Replace(CodBeneficiario, "'", "''") & "' "
    End If

    If Trim$(GrupoDesde) <> "" Then FiltroGrupos = FiltroGrupos & _
        "AND Cuentas.KeyGrupo >= '" & Replace(Trim$(GrupoDesde), "'", "''") & "' "
    If Trim$(GrupoHasta) <> "" Then FiltroGrupos = FiltroGrupos & _
        "AND Cuentas.KeyGrupo <= '" & Replace(Trim$(GrupoHasta), "'", "''") & "' "

    '--- Piezas repetidas
    JoinIndice = "INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion " & _
                 "AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo " & _
                 "AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento "
    JoinTasas = "INNER JOIN Tasas ON IndiceTransaccion.FechaTransaccion = Tasas.FechaTasas "

    ExprDebito = "CASE WHEN IndiceTransaccion.TipoMoneda = '" & MonedaSql & "' THEN Transacciones.Debito " & _
                 "ELSE Transacciones.Debito" & OperadorTasa & "Tasas.MontoCordobas END"
    ExprCredito = "CASE WHEN IndiceTransaccion.TipoMoneda = '" & MonedaSql & "' THEN Transacciones.Credito " & _
                  "ELSE Transacciones.Credito" & OperadorTasa & "Tasas.MontoCordobas END"

    '--- SELECT común
    Sql = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, "
    Sql = Sql & "Cuentas.KeyGrupo, Grupos.DescripcionGrupo, "
    Sql = Sql & "ISNULL(S.DebitoAnterior, 0) AS DebitoAnterior, "
    Sql = Sql & "ISNULL(S.CreditoAnterior, 0) AS CreditoAnterior, "
    Sql = Sql & "M.CodBeneficiario, M.FechaTransaccion, M.NPeriodo, M.NTransaccion, "
    Sql = Sql & "M.NumeroMovimiento, M.VoucherNo, M.DescripcionMovimiento, M.Clave, "
    Sql = Sql & "M.FacturaNo, M.ChequeNo, M.Fuente, M.TCambio, M.Debito, M.Credito "
    Sql = Sql & "FROM Cuentas "

    '--- A: cuentas con acumulado hasta FechaFin
    Sql = Sql & "INNER JOIN (SELECT DISTINCT Transacciones.CodCuentas FROM Transacciones " & JoinIndice
    Sql = Sql & "WHERE Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & F2 & "', 102) "
    Sql = Sql & "AND IndiceTransaccion.Ajuste <> '" & AjusteExcluir & "' "
    Sql = Sql & "AND Transacciones.Debito + Transacciones.Credito <> 0 "
    Sql = Sql & FiltroCuentas & FiltroBenef
    Sql = Sql & ") AS A ON A.CodCuentas = Cuentas.CodCuentas "

    '--- Grupo de la cuenta
    Sql = Sql & "LEFT JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo "

    '--- S: acumulado anterior a FechaIni (saldo inicial)
    Sql = Sql & "LEFT JOIN (SELECT Transacciones.CodCuentas, "
    Sql = Sql & "SUM(" & ExprDebito & ") AS DebitoAnterior, "
    Sql = Sql & "SUM(" & ExprCredito & ") AS CreditoAnterior "
    Sql = Sql & "FROM Transacciones " & JoinIndice & JoinTasas
    Sql = Sql & "WHERE Transacciones.FechaTransaccion < CONVERT(DATETIME, '" & F1 & "', 102) "
    Sql = Sql & "AND IndiceTransaccion.Ajuste <> '" & AjusteExcluir & "' "
    Sql = Sql & FiltroCuentas & FiltroBenef
    Sql = Sql & "GROUP BY Transacciones.CodCuentas"
    Sql = Sql & ") AS S ON S.CodCuentas = Cuentas.CodCuentas "

    '--- M: movimientos del periodo (detallados o totalizados)
    Sql = Sql & "LEFT JOIN (SELECT Transacciones.CodCuentas, "

    If Detallado Then
        Sql = Sql & "Transacciones.Cod_Tercero AS CodBeneficiario, "
        Sql = Sql & "Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, "
        Sql = Sql & "Transacciones.NumeroMovimiento, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, "
        Sql = Sql & "Transacciones.Clave, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Fuente, "
        Sql = Sql & "Transacciones.TCambio, "
        Sql = Sql & ExprDebito & " AS Debito, " & ExprCredito & " AS Credito "
    Else
        If AgruparPorBeneficiario Then
            Sql = Sql & "Transacciones.Cod_Tercero AS CodBeneficiario, "
        Else
            Sql = Sql & "CAST(NULL AS VARCHAR(50)) AS CodBeneficiario, "
        End If
        Sql = Sql & "MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, "
        Sql = Sql & "MAX(Transacciones.NPeriodo) AS NPeriodo, "
        Sql = Sql & "COUNT(*) AS NTransaccion, "
        Sql = Sql & "MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, "
        Sql = Sql & "NULL AS VoucherNo, 'MOVIMIENTOS DEL PERIODO' AS DescripcionMovimiento, "
        Sql = Sql & "NULL AS Clave, NULL AS FacturaNo, NULL AS ChequeNo, NULL AS Fuente, NULL AS TCambio, "
        Sql = Sql & "SUM(" & ExprDebito & ") AS Debito, SUM(" & ExprCredito & ") AS Credito "
    End If

    Sql = Sql & "FROM Transacciones " & JoinIndice & JoinTasas
    Sql = Sql & "WHERE Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & F1 & "', 102) "
    Sql = Sql & "AND CONVERT(DATETIME, '" & F2 & "', 102) "
    Sql = Sql & "AND IndiceTransaccion.Ajuste <> '" & AjusteExcluir & "' "
    Sql = Sql & FiltroCuentas & FiltroBenef

    If Not Detallado Then
        If AgruparPorBeneficiario Then
            Sql = Sql & "GROUP BY Transacciones.CodCuentas, Transacciones.Cod_Tercero "
        Else
            Sql = Sql & "GROUP BY Transacciones.CodCuentas "
        End If
    End If

    Sql = Sql & ") AS M ON M.CodCuentas = Cuentas.CodCuentas "

    '--- Filtro por grupos
    If FiltroGrupos <> "" Then Sql = Sql & "WHERE 1 = 1 " & FiltroGrupos

    '--- Orden
    If AgruparPorBeneficiario Then
        Sql = Sql & "ORDER BY Cuentas.CodCuentas, M.CodBeneficiario, M.FechaTransaccion, M.NTransaccion"
    Else
        Sql = Sql & "ORDER BY Cuentas.CodCuentas, M.FechaTransaccion, M.NTransaccion"
    End If

    SqlAuxiliarCuentas = Sql

End Function

'==========================================================
' Naturaleza de la cuenta (deudora = Debe - Haber)
'==========================================================
Public Function EsCuentaDeudora(ByVal Tipo As String) As Boolean
    Select Case Tipo
        Case "Activo Fijo", "Otros Activos", "Caja", "Cuentas x Cobrar", "Bancos", _
             "Costos", "Gastos", "Papeleria - Utiles", "Inventario"
            EsCuentaDeudora = True
    End Select
End Function

'==========================================================
' Null a 0
'==========================================================
Public Function NzD(ByVal Valor As Variant) As Double
    If IsNull(Valor) Or IsEmpty(Valor) Then NzD = 0 Else NzD = CDbl(Valor)
End Function



Public Function ObtenerRetenciones( _
    ByVal FechaDesde As Date, _
    ByVal FechaHasta As Date, _
    ByVal CodigoCuenta As String) As ADODB.Recordset

    Dim Rs As ADODB.Recordset
    Dim Sql As String

    On Error GoTo ErrorHandler

    Set Rs = New ADODB.Recordset

    '------------------------------------------------------------
    ' CONSULTA DE RETENCIONES
    '------------------------------------------------------------
    Sql = ""
    Sql = Sql & "SELECT "
    
    Sql = Sql & "T.FechaTransaccion, "
    Sql = Sql & "T.NPeriodo, "
    Sql = Sql & "T.NumeroMovimiento, "
    Sql = Sql & "T.NTransaccion, "
    
    Sql = Sql & "T.CodCuentas, "
    Sql = Sql & "C.DescripcionCuentas AS NombreCuenta, "
    Sql = Sql & "C.DescRetencion, "
    
    Sql = Sql & "T.Cod_Tercero, "
    Sql = Sql & "T.Beneficiario, "
    
    Sql = Sql & "T.FacturaNo, "
    Sql = Sql & "T.NumNomina, "
    Sql = Sql & "T.MontoBase, "
    
    Sql = Sql & "T.DescripcionMovimiento, "
    Sql = Sql & "T.VoucherNo, "
    Sql = Sql & "T.ChequeNo, "
    Sql = Sql & "T.Fuente, "
    
    Sql = Sql & "ISNULL(T.Debito,0) AS Debito, "
    Sql = Sql & "ISNULL(T.Credito,0) AS Credito, "
    
    '------------------------------------------------------------
    ' RETENCION REALMENTE CONTABILIZADA
    '
    ' Cuenta de pasivo:
    '     Credito = aumenta la retencion
    '     Debito  = disminuye/anula la retencion
    '
    ' Por tanto:
    '     Retenido = Credito - Debito
    '------------------------------------------------------------
    Sql = Sql & "(ISNULL(T.Credito,0) - ISNULL(T.Debito,0)) AS Retenido "

    Sql = Sql & "FROM Transacciones AS T "

    Sql = Sql & "INNER JOIN Cuentas AS C "
    Sql = Sql & "ON T.CodCuentas = C.CodCuentas "

    Sql = Sql & "WHERE "

    '------------------------------------------------------------
    ' CUENTA SOLICITADA
    '------------------------------------------------------------
    Sql = Sql & "T.CodCuentas = '" & _
                Replace(Trim$(CodigoCuenta), "'", "''") & "' "

    '------------------------------------------------------------
    ' LA CUENTA DEBE SER DE RETENCION
    '------------------------------------------------------------
    Sql = Sql & "AND ISNULL(C.CausaRetencion,0) = 1 "

    '------------------------------------------------------------
    ' FECHA DESDE
    '------------------------------------------------------------
    Sql = Sql & "AND T.FechaTransaccion >= '" & _
                Format$(FechaDesde, "yyyymmdd") & "' "

    '------------------------------------------------------------
    ' FECHA HASTA
    '
    ' Utilizamos < fecha siguiente para incluir todo el día
    ' aunque FechaTransaccion tenga componente de hora.
    '------------------------------------------------------------
    Sql = Sql & "AND T.FechaTransaccion < '" & _
                Format$(DateAdd("d", 1, FechaHasta), "yyyymmdd") & "' "

    '------------------------------------------------------------
    ' ORDEN
    '------------------------------------------------------------

    Sql = Sql & "ORDER BY "
    Sql = Sql & "T.CodCuentas, "
    Sql = Sql & "T.FechaTransaccion, "
    Sql = Sql & "T.NumeroMovimiento, "
    Sql = Sql & "T.NTransaccion "

    '------------------------------------------------------------
    ' EJECUTAR
    '------------------------------------------------------------
    With Rs

        .CursorLocation = adUseClient

        .Open Sql, _
              ConexionReporte, _
              adOpenStatic, _
              adLockReadOnly

    End With

    Set ObtenerRetenciones = Rs

    Exit Function

ErrorHandler:

    If Not Rs Is Nothing Then

        If Rs.State <> 0 Then
            Rs.Close
        End If

    End If

    Set Rs = Nothing

    MsgBox "Error al obtener las retenciones." & vbCrLf & _
           "Número: " & err.Number & vbCrLf & _
           "Descripción: " & err.Description, _
           vbExclamation, _
           "Reporte de Retenciones"

End Function


Private Function FormatoFechaSQL(ByVal Fecha As Date) As String

    FormatoFechaSQL = _
        "'" & Format$(Fecha, "yyyyMMdd") & "'"

End Function


 
     
     
     
     
     
     
     
     
     
     
     
     
'  Case "AUXILIAR DE CUENTAS"
'
'    NumFecha1 = Me.DTFecha1.Value
'    NumFecha2 = Me.DTFecha2.Value
'
'
''         If FrmReportes.CmbMoneda.Text = "Córdobas" Then
''            Ajuste = "Dólares"
''         ElseIf FrmReportes.CmbMoneda.Text = "Dólares" Then
''            Ajuste = "Córdobas"
''
''         End If
'
'        Dim FiltroBeneficiario As String
'
'        FiltroBeneficiario = ""
'
'        If ActivarBeneficiario = 1 Then
'
'            If CodigoBeneficiarioSeleccionado <> "" Then
'
'                FiltroBeneficiario = _
'                    " AND Transacciones.Cod_Tercero = '" & _
'                    Replace(CodigoBeneficiarioSeleccionado, "'", "''") & "'"
'
'            End If
'
'        End If
'
'
'      If Me.Option4.Value = True Then
'          ArepAuxiliar.LblRangoFecha = "Desde " & Me.DTFecha1.Value & " Hasta " & Me.DTFecha2.Value
'          ArepAuxiliar.LblFecha = Format(Now, "dd/mm/yyyy")
'          ArepAuxiliar.DataControl1.ConnectionString = ConexionReporte
'
'
'            If Me.DBCodigo.Text = "" Then
'                 Me.DtaConsulta.RecordSource = "SELECT Cuentas.* From Cuentas ORDER BY CodCuentas"
'                 Me.DtaConsulta.Refresh
'                 If Not Me.DtaConsulta.Recordset.EOF Then
'                   Me.DtaConsulta.Recordset.MoveFirst
'                   CodigoCuentaDesde = Me.DtaConsulta.Recordset("CodCuentas")
'                End If
'            Else
'              CodigoCuentaDesde = Me.DBCodigo.Text
'            End If
'
'            If Me.DBCodigoHasta.Text = "" Then
'                 Me.DtaConsulta.RecordSource = "SELECT Cuentas.* From Cuentas ORDER BY CodCuentas"
'                 Me.DtaConsulta.Refresh
'                 If Not Me.DtaConsulta.Recordset.EOF Then
'                   Me.DtaConsulta.Recordset.MoveLast
'                   CodigoCuentaHasta = Me.DtaConsulta.Recordset("CodCuentas")
'                End If
'            Else
'               CodigoCuentaHasta = Me.DBCodigoHasta.Text
'            End If
'
'
'            CodDesde = CodigoCuentaDesde
'            CodHasta = CodigoCuentaHasta
'
'             If Me.CmbMoneda.Text = "Córdobas" Then
'
'               Moneda = "Cordobas"
'               Sql = "SELECT Transacciones.CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Transacciones.NPeriodo) AS NPeriodo, MAX(Transacciones.NTransaccion) AS NTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.VoucherNo) AS VoucherNo, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.Clave) AS Clave, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Debito ELSE Transacciones.Debito * Tasas.MontoCordobas END) AS Debito, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Córdobas' THEN Transacciones.Credito ELSE Transacciones.Credito * Tasas.MontoCordobas END) AS Credito, MAX(Transacciones.FacturaNo) AS FacturaNo, MAX(Transacciones.ChequeNo) AS ChequeNo, MAX(Transacciones.Fuente) AS Fuente, MAX(Cuentas.TipoCuenta) AS TipoCuenta, SUM(Transacciones.Debito + Transacciones.Credito) AS Saldo FROM  Cuentas INNER JOIN " & _
'                     "Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN Tasas ON IndiceTransaccion.FechaTransaccion = Tasas.FechaTasas  " & _
'                     "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) AND (IndiceTransaccion.Ajuste <> 'Dólares') GROUP BY Transacciones.CodCuentas HAVING (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') AND (SUM(Transacciones.Debito + Transacciones.Credito) <> 0) ORDER BY Transacciones.CodCuentas"
'             Else
'               Moneda = "Dolares"
'               Sql = "SELECT Transacciones.CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Transacciones.NPeriodo) AS NPeriodo, MAX(Transacciones.NTransaccion) AS NTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.VoucherNo) AS VoucherNo, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.Clave) AS Clave, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Tasas.MontoCordobas END) AS Debito, SUM(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Tasas.MontoCordobas END) AS Credito, MAX(Transacciones.FacturaNo) AS FacturaNo, MAX(Transacciones.ChequeNo) AS ChequeNo, MAX(Transacciones.Fuente) AS Fuente, MAX(Cuentas.TipoCuenta) AS TipoCuenta, SUM(Transacciones.Debito + Transacciones.Credito) AS Saldo FROM  Cuentas INNER JOIN " & _
'                     "Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento INNER JOIN Tasas ON IndiceTransaccion.FechaTransaccion = Tasas.FechaTasas  " & _
'                     "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) AND (IndiceTransaccion.Ajuste <> 'Córdobas') GROUP BY Transacciones.CodCuentas HAVING (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') AND (SUM(Transacciones.Debito + Transacciones.Credito) <> 0) ORDER BY Transacciones.CodCuentas"
'
'             End If
'
''            SQL = "SELECT Transacciones.CodCuentas,  MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Transacciones.NPeriodo) AS NPeriodo,MAX(Transacciones.NTransaccion) AS NTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.VoucherNo) AS VoucherNo, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.Clave) AS Clave, SUM(Transacciones.Debito) AS Debito, SUM(Transacciones.Credito) AS Credito, MAX(Transacciones.FacturaNo) AS FacturaNo, MAX(Transacciones.ChequeNo) AS ChequeNo, MAX(Transacciones.Fuente) AS Fuente, MAX(Cuentas.TipoCuenta) AS TipoCuenta, SUM(Transacciones.Debito + Transacciones.Credito) As Saldo FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE  (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas " & _
''                                               "HAVING (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') AND (SUM(Transacciones.Debito + Transacciones.Credito) <> 0) ORDER BY Transacciones.CodCuentas"
'
'               ArepAuxiliar.DataControl1.ConnectionString = ConexionReporte
'
''             ArepAuxiliar.LblCodigo.Caption = "Mis Code:" & Me.DBCodigo.Text
''             ArepAuxiliar.LblRango.Caption = "Filtrado Desde: " & CodigoCuentaDesde & " Hasta " & CodigoCuentaHasta
''             If Dir(RutaLogo) <> "" Then
''                   ArepAuxiliar.Logo.Picture = LoadPicture(RutaLogo)
''             End If
''             ArepAuxiliar.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
''             ArepAuxiliar.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
''             ArepAuxiliar.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
''             ArepAuxiliar.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
'
'            Set rpt = New ArepAuxiliar
'            rpt.DataControl1.ConnectionString = ConexionReporte
'            rpt.DataControl1.Source = Sql
'            fPreview.RunReport rpt
'            fPreview.Show 1
''             ArepAuxiliar.Show 1
''            fPreview.arv.ReportSource = ArepAuxiliar
''            fPreview.Show 1
'
'       ElseIf Me.Option5.Value = True Then
'
'          ArepAuxiliarMayor.LblRangoFecha = "Desde " & Me.DTFecha1.Value & " Hasta " & Me.DTFecha2.Value
'          ArepAuxiliarMayor.LblFecha = Format(Now, "dd/mm/yyyy")
'          ArepAuxiliarMayor.DataControl1.ConnectionString = ConexionReporte
'
'            If Me.TxtDesde.Text = "" Then
'               Me.DtaConsulta.RecordSource = "SELECT * From Grupos ORDER BY KeyGrupo"
'               Me.DtaConsulta.Refresh
'               If Not Me.DtaConsulta.Recordset.EOF Then
'                 Me.DtaConsulta.Recordset.MoveFirst
'                 CodigoCuentaDesde = Me.DtaConsulta.Recordset("KeyGrupo")
'               End If
'            Else
'                CodigoCuentaDesde = Me.TxtKeyGrupoDesde.Text
'            End If
'
'            If Me.TxtHasta.Text = "" Then
'               Me.DtaConsulta.RecordSource = "SELECT * From Grupos ORDER BY KeyGrupo"
'               Me.DtaConsulta.Refresh
'               If Not Me.DtaConsulta.Recordset.EOF Then
'                 Me.DtaConsulta.Recordset.MoveLast
'                 CodigoCuentaHasta = Me.DtaConsulta.Recordset("KeyGrupo")
'               End If
'            Else
'               CodigoCuentaHasta = Me.TxtKeyGrupoHasta.Text
'            End If
'
'             Sql = "SELECT  Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.TCambio, Transacciones.FacturaNo,Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTransaccion, Cuentas.TipoCuenta, Transacciones.TCambio AS Expr1, Transacciones.TCambio * Transacciones.Debito AS Debito, Transacciones.TCambio * Transacciones.Credito AS Credito, Grupos.KeyGrupo, Grupos.DescripcionGrupo " & _
'                    "FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas INNER JOIN Grupos ON Cuentas.KeyGrupo = Grupos.KeyGrupo " & _
'                    "WHERE  (Transacciones.FechaTransaccion BETWEEN '" & Format(Me.DTFecha1, "yyyymmdd") & "' AND '" & Format(Me.DTFecha2, "yyyymmdd") & "') AND (Grupos.KeyGrupo BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') " & _
'                    "ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion "
'
'             ArepAuxiliarMayor.DataControl1.Source = Sql
'             ArepAuxiliarMayor.LblCodigo.Caption = "Mis Code:" & Me.DBCodigo.Text
'             ArepAuxiliarMayor.LblRango.Caption = "Filtrado Desde: " & CodigoCuentaDesde & " Hasta " & CodigoCuentaHasta
'             ArepAuxiliarMayor.Logo.Picture = LoadPicture(RutaLogo)
'             ArepAuxiliarMayor.LblEmpresa = Me.DtaDatosEmpresa.Recordset("NombreEmpresa")
'             ArepAuxiliarMayor.LblEmpresa1 = Me.DtaDatosEmpresa.Recordset("Direccion")
'             ArepAuxiliarMayor.LblEmpresa2 = "RUC: " & Me.DtaDatosEmpresa.Recordset("NumeroRuc")
'             ArepAuxiliarMayor.LblFechaImpreso = Format(Now, "dd/mm/yyyy")
''             ArepAuxiliarMayor.Show 1
'
'            Set rpt = New ArepAuxiliar
'            rpt.DataControl1.ConnectionString = ConexionReporte
'            rpt.DataControl1.Source = Sql
'            fPreview.RunReport rpt
'            fPreview.Show 1
'
'       End If
'


'==============================================================================
' Retenciones por Beneficiario
' Une las retenciones de la contabilidad (Conexion) con el nombre del
' beneficiario del Sistema de Facturacion (ConexionFacturacion).
' Devuelve un recordset desconectado listo para DataControl1.Recordset.
' Campos: Fecha, Documento, Cod_Tercero, NombreBeneficiario, Retenido,
'         MontoBase, NTransaccion (los dos ultimos solo por compatibilidad
'         con los controles de resumen del layout base).
'==============================================================================
Public Function ObtenerRsRetencionesBeneficiario( _
    ByVal FechaDesde As Date, _
    ByVal FechaHasta As Date, _
    ByVal CuentaDesde As String, _
    ByVal CuentaHasta As String, _
    Optional ByVal Resumido As Boolean = False) As ADODB.Recordset

    Dim CnConta As ADODB.Connection
    Dim CnFact As ADODB.Connection
    Dim RsConta As ADODB.Recordset
    Dim RsFact As ADODB.Recordset
    Dim RsSal As ADODB.Recordset
    Dim RsRes As ADODB.Recordset
    Dim CodAnt As String
    Dim NomAnt As String
    Dim Acum As Double
    Dim Hay As Boolean
    Dim Nombres As Object
    Dim Cedulas As Object
    Dim Cedula As String
    Dim Sql As String
    Dim Cod As String
    Dim Nombre As String
    Dim Fecha As Date
    Dim Monto As Double
    Dim NumErr As Long
    Dim DescErr As String

    On Error GoTo ErrFn

    '--- Nombres de beneficiarios (Sistema de Facturacion)
    Set Nombres = CreateObject("Scripting.Dictionary")
    Nombres.CompareMode = 1 'vbTextCompare
    Set Cedulas = CreateObject("Scripting.Dictionary")
    Cedulas.CompareMode = 1

    Set CnFact = New ADODB.Connection
    CnFact.Open ConexionFacturacion

    Set RsFact = CnFact.Execute("SELECT Codigo_Beneficiario, " & _
        "LTRIM(RTRIM(ISNULL(Nombre_Beneficiario,'') + ' ' + ISNULL(Apellido_Beneficiario,''))) AS Nombre, " & _
        "LTRIM(RTRIM(ISNULL(NULLIF(LTRIM(RTRIM(Cedula)), ''), ISNULL(RUC, '')))) AS Cedula " & _
        "FROM Beneficiario")

    Do While Not RsFact.EOF
        Cod = Trim$(RsFact!Codigo_Beneficiario & "")
        If Not Nombres.Exists(Cod) Then
            Nombres.Add Cod, Trim$(RsFact!Nombre & "")
            Cedulas.Add Cod, Trim$(RsFact!Cedula & "")
        End If
        RsFact.MoveNext
    Loop

    RsFact.Close
    CnFact.Close

    '--- Retenciones (Contabilidad): Credito - Debito por linea de transaccion
    Sql = "SELECT T.FechaTransaccion, T.NumeroMovimiento, T.NTransaccion, " & _
          "ISNULL(I.Fuente, ISNULL(T.Fuente, '')) AS Fuente, " & _
          "LTRIM(RTRIM(ISNULL(T.Cod_Tercero, ''))) AS Cod_Tercero, " & _
          "(ISNULL(T.Credito, 0) - ISNULL(T.Debito, 0)) AS Retenido, " & _
          "ISNULL(CONVERT(VARCHAR(30), C.DescRetencion), '') AS DescRetencion " & _
          "FROM Transacciones AS T " & _
          "INNER JOIN Cuentas AS C ON T.CodCuentas = C.CodCuentas " & _
          "LEFT JOIN IndiceTransaccion AS I ON T.FechaTransaccion = I.FechaTransaccion " & _
          "AND T.NPeriodo = I.NPeriodo AND T.NumeroMovimiento = I.NumeroMovimiento " & _
          "WHERE ISNULL(C.CausaRetencion, 0) = 1 " & _
          "AND (ISNULL(T.Credito, 0) - ISNULL(T.Debito, 0)) <> 0 "

    If Len(Trim$(CuentaDesde)) > 0 Then
        Sql = Sql & "AND T.CodCuentas >= '" & Replace(Trim$(CuentaDesde), "'", "''") & "' "
    End If
    If Len(Trim$(CuentaHasta)) > 0 Then
        Sql = Sql & "AND T.CodCuentas <= '" & Replace(Trim$(CuentaHasta), "'", "''") & "' "
    End If

    Sql = Sql & "AND T.FechaTransaccion >= '" & Format$(FechaDesde, "yyyymmdd") & "' "
    Sql = Sql & "AND T.FechaTransaccion < '" & Format$(DateAdd("d", 1, FechaHasta), "yyyymmdd") & "' "
    Sql = Sql & "ORDER BY T.FechaTransaccion, T.NumeroMovimiento, T.NTransaccion"

    Set CnConta = New ADODB.Connection
    CnConta.Open Conexion
    Set RsConta = CnConta.Execute(Sql)

    '--- Recordset de salida (desconectado)
    Set RsSal = New ADODB.Recordset
    RsSal.CursorLocation = adUseClient
    AgregarCamposRetBenef RsSal
    RsSal.Open


    Do While Not RsConta.EOF
        Fecha = RsConta!FechaTransaccion
        Cod = Trim$(RsConta!Cod_Tercero & "")
        Monto = CDbl(RsConta!Retenido)

        Cedula = ""
        If Cedulas.Exists(Cod) Then Cedula = Cedulas(Cod)

        If Len(Cod) = 0 Then
            Nombre = "SIN BENEFICIARIO"
        ElseIf Nombres.Exists(Cod) Then
            Nombre = Nombres(Cod)
        Else
            Nombre = "BENEFICIARIO NO ENCONTRADO"
        End If

        RsSal.AddNew
        RsSal!Fecha = Fecha
        ' Fuente (primeros 4) + ddmmyyyy + "000" + NumeroMovimiento
        RsSal!Documento = Left$(Trim$(RsConta!Fuente & ""), 4) & Format$(Fecha, "ddmmyyyy") & _
                          "000" & Trim$(RsConta!NumeroMovimiento & "")
        RsSal!Cod_Tercero = Cod
        RsSal!NombreBeneficiario = Nombre
        RsSal!Retenido = Monto
        RsSal!MontoBase = Monto
        RsSal!NTransaccion = CDbl(Val(RsConta!NTransaccion & ""))
        RsSal!Grupo = 1
        RsSal!Cedula = Cedula
        If Len(Trim$(RsConta!DescRetencion & "")) > 0 Then
            RsSal!TasaRet = Trim$(CStr(Val(RsConta!DescRetencion))) & " %"
        End If
        RsSal.Update

        RsConta.MoveNext
    Loop

    RsConta.Close
    CnConta.Close

    If RsSal.RecordCount > 0 Then
        RsSal.Sort = "Cod_Tercero ASC, Fecha ASC, Documento ASC"
        RsSal.MoveFirst
    End If

    '--- Resumido: un registro por beneficiario con el total retenido
    If Resumido Then
        Set RsRes = New ADODB.Recordset
        RsRes.CursorLocation = adUseClient
        AgregarCamposRetBenef RsRes
        RsRes.Open

        Hay = False
        Acum = 0
        Do While Not RsSal.EOF
            If Hay Then
                If RsSal!Cod_Tercero <> CodAnt Then
                    AgregarFilaRetBenef RsRes, CodAnt, NomAnt, Acum
                    Acum = 0
                End If
            End If
            CodAnt = RsSal!Cod_Tercero & ""
            NomAnt = RsSal!NombreBeneficiario & ""
            Acum = Acum + RsSal!Retenido
            Hay = True
            RsSal.MoveNext
        Loop
        If Hay Then AgregarFilaRetBenef RsRes, CodAnt, NomAnt, Acum

        RsSal.Close
        Set RsSal = RsRes
        If RsSal.RecordCount > 0 Then
            RsSal.Sort = "NombreBeneficiario ASC, Cod_Tercero ASC"
            RsSal.MoveFirst
        End If
    End If

    Set ObtenerRsRetencionesBeneficiario = RsSal
    Exit Function

ErrFn:
    NumErr = err.Number
    DescErr = err.Description
    On Error Resume Next
    If Not RsFact Is Nothing Then If RsFact.State = adStateOpen Then RsFact.Close
    If Not RsConta Is Nothing Then If RsConta.State = adStateOpen Then RsConta.Close
    If Not CnFact Is Nothing Then If CnFact.State = adStateOpen Then CnFact.Close
    If Not CnConta Is Nothing Then If CnConta.State = adStateOpen Then CnConta.Close
    err.Raise NumErr, "ObtenerRsRetencionesBeneficiario", DescErr

End Function

Private Sub AgregarCamposRetBenef(ByRef Rs As ADODB.Recordset)
    Rs.Fields.Append "Fecha", adDate
    Rs.Fields.Append "Documento", adVarWChar, 30
    Rs.Fields.Append "Cod_Tercero", adVarWChar, 30
    Rs.Fields.Append "NombreBeneficiario", adVarWChar, 200
    Rs.Fields.Append "Retenido", adDouble
    Rs.Fields.Append "MontoBase", adDouble
    Rs.Fields.Append "NTransaccion", adDouble
    Rs.Fields.Append "Grupo", adDouble
    ' Campos vacios: los controles ocultos del layout base siguen enlazados a ellos
    Rs.Fields.Append "Cedula", adVarWChar, 30
    Rs.Fields.Append "TasaRet", adVarWChar, 20
    Rs.Fields.Append "ChequeNo", adVarWChar, 5
    Rs.Fields.Append "DescripcionMovimiento", adVarWChar, 5
    Rs.Fields.Append "DescRetencion", adVarWChar, 5
End Sub

Private Sub AgregarFilaRetBenef(ByRef Rs As ADODB.Recordset, ByVal Cod As String, _
                                ByVal Nombre As String, ByVal Monto As Double)
    Rs.AddNew
    Rs!Cod_Tercero = Cod
    Rs!NombreBeneficiario = Nombre
    Rs!Retenido = Monto
    Rs!MontoBase = Monto
    Rs!NTransaccion = 0
    Rs!Grupo = 1
    Rs.Update
End Sub