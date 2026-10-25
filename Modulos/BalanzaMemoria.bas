Attribute VB_Name = "BalanzaMemoria"
Option Explicit
'=====================================================================================================
' BALANZA DE COMPROBACION - VERSION EN MEMORIA
'
' Equivale a SaldoReportes("BalanzaCodigo") + el "parche" de FrmReportes (debe2/haber2 y debe3/haber3),
' pero sin tocar la tabla Reportes: una sola consulta trae, por cuenta, lo anterior a FechaDesde y la
' actividad FechaDesde..FechaHasta, y las columnas se arman aqui:
'   Debe1/Haber1  saldo anterior, en la columna natural de la cuenta (0 si SoloPeriodo)
'   Debe2/Haber2  debitos y creditos del periodo
'   Debe3/Haber3  saldo final: la diferencia (Debe1+Debe2) - (Haber1+Haber2) en la columna que le toque
' Mientras RptBalanzaActivo = True, ArepBalanza no consulta por cuenta (Detail_Format) y calcula el pie
' desde esta memoria. Siempre llamar RptBalanzaLiberar al terminar.
'=====================================================================================================

Public RptBalanzaActivo As Boolean

Private mRsBal As ADODB.Recordset

Private Function Dbl(ByVal v As Variant) As Double
    If IsNull(v) Then
        Dbl = 0
    ElseIf IsEmpty(v) Then
        Dbl = 0
    Else
        Dbl = CDbl(v)
    End If
End Function

Private Function SqlFecha(ByVal Fecha As Date) As String
    SqlFecha = Format$(Int(Fecha), "yyyymmdd")
End Function

Private Function SqlTexto(ByVal Texto As String) As String
    SqlTexto = Replace(Texto, "'", "''")
End Function

' Naturaleza deudora segun ArepBalanza; el resto de los tipos es acreedor.
Private Function EsDeudora(ByVal Tipo As String) As Boolean
    Select Case Tipo
        Case "Activo Fijo", "Otros Activos", "Caja", "Cuentas x Cobrar", "Bancos", "Costos", "Gastos", _
             "Papeleria - Utiles", "Inventario"
            EsDeudora = True
    End Select
End Function

Private Function NuevoRecordset() As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    With rs.Fields
        .Append "Descripcion", adVarWChar, 255, adFldIsNullable
        .Append "Debe1", adDouble
        .Append "Haber1", adDouble
        .Append "Debe2", adDouble
        .Append "Haber2", adDouble
        .Append "Debe3", adDouble
        .Append "Haber3", adDouble
        .Append "TotalDebe", adDouble
        .Append "TotalHaber", adDouble
        .Append "KeyGrupo", adVarWChar, 50, adFldIsNullable
        .Append "KeyGrupoSuperior", adVarWChar, 50, adFldIsNullable
        .Append "KeyGrupoCuenta", adVarWChar, 50, adFldIsNullable
        .Append "Nivel", adInteger
        .Append "Orden", adInteger
        .Append "CodDepartamento", adVarWChar, 50, adFldIsNullable
        .Append "CodCuentas", adVarWChar, 50, adFldIsNullable
        .Append "Ubicacion", adVarWChar, 100, adFldIsNullable
        .Append "TipoCuenta", adVarWChar, 50, adFldIsNullable
        .Append "DescripcionCuentas", adVarWChar, 255, adFldIsNullable
    End With
    rs.Open
    Set NuevoRecordset = rs
End Function

' Falla si hay movimientos hasta FechaHasta sin tasa de cambio (el original abortaba igual).
Private Sub ValidaTasas(cn As ADODB.Connection, ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal SoloPeriodo As Boolean)
    Dim rs As ADODB.Recordset, Sql As String
    Sql = "SELECT COUNT(*) FROM Transacciones t LEFT JOIN Tasas ta ON t.FechaTransaccion = ta.FechaTasas " & _
          "WHERE t.FechaTransaccion < '" & SqlFecha(DateAdd("d", 1, Int(FechaHasta))) & "' "
    If SoloPeriodo Then Sql = Sql & "AND t.FechaTransaccion >= '" & SqlFecha(FechaDesde) & "' "
    Sql = Sql & "AND (ta.MontoCordobas IS NULL OR ta.MontoCordobas = 0)"
    Set rs = cn.Execute(Sql)
    If CLng(rs.Fields(0).Value) > 0 Then
        rs.Close
        Err.Raise vbObjectError + 2001, "BalanzaMemoria", _
            "Existen movimientos sin Tasa de Cambio (o con tasa cero)." & vbLf & _
            "El reporte en Dolares no puede continuar."
    End If
    rs.Close
End Sub

'-----------------------------------------------------------------------------------------------------
' RptBalanzaCodigoGenerar  (BALANZA DE COMPROBACION - Por Codigo)
'   SoloPeriodo  = casilla "Solo Historial Periodo Seleccionado" (sin saldo anterior)
'   CodDesde/CodHasta  rango de cuentas (vacio = todas)
'   ExcluirCierre  quita del periodo los asientos de Fuente = 'Cierre' (casilla Quitar Movimientos de Cierre)
' Siempre se excluyen los asientos de ajuste de la otra moneda, como hace el original.
'-----------------------------------------------------------------------------------------------------
Public Function RptBalanzaCodigoGenerar(ByVal FechaDesde As Date, ByVal FechaHasta As Date, _
        ByVal EnDolares As Boolean, ByVal ExcluirCierre As Boolean, ByVal SoloPeriodo As Boolean, _
        ByVal CodDesde As String, ByVal CodHasta As String) As ADODB.Recordset
    Dim cn As ADODB.Connection, rsSql As ADODB.Recordset, rs As ADODB.Recordset, Sql As String
    Dim Ini As String, Debito As String, Credito As String, CondPeriodo As String, Ajuste As String
    Dim D1 As Double, C1 As Double, D2 As Double, C2 As Double, Net As Double
    Dim Debe1 As Double, Haber1 As Double, Tipo As String, Desc As String, Orden As Long
    Dim ErrNum As Long, ErrDesc As String, ErrSrc As String

    RptBalanzaLiberar
    If Len(CodHasta) = 0 Then CodHasta = "zzzzzzzzzzzzzzzz"
    If EnDolares Then Ajuste = "C" & ChrW(243) & "rdobas" Else Ajuste = "D" & ChrW(243) & "lares"
    Ini = "'" & SqlFecha(FechaDesde) & "'"
    If EnDolares Then
        Debito = "t.Debito * t.TCambio / ta.MontoCordobas"
        Credito = "t.Credito * t.TCambio / ta.MontoCordobas"
    Else
        Debito = "t.Debito * t.TCambio"
        Credito = "t.Credito * t.TCambio"
    End If
    CondPeriodo = "t.FechaTransaccion >= " & Ini
    If ExcluirCierre Then CondPeriodo = CondPeriodo & " AND ISNULL(i.Fuente, '') <> 'Cierre'"

    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open Conexion
    On Error GoTo Falla
    If EnDolares Then ValidaTasas cn, FechaDesde, FechaHasta, SoloPeriodo

    Sql = "SELECT c.CodCuentas, c.DescripcionCuentas, c.TipoCuenta, " & _
          "SUM(CASE WHEN t.FechaTransaccion < " & Ini & " THEN " & Debito & " END) AS D1, " & _
          "SUM(CASE WHEN t.FechaTransaccion < " & Ini & " THEN " & Credito & " END) AS C1, " & _
          "SUM(CASE WHEN " & CondPeriodo & " THEN " & Debito & " END) AS D2, " & _
          "SUM(CASE WHEN " & CondPeriodo & " THEN " & Credito & " END) AS C2 " & _
          "FROM Cuentas c INNER JOIN Transacciones t ON c.CodCuentas = t.CodCuentas " & _
          "INNER JOIN IndiceTransaccion i ON t.FechaTransaccion = i.FechaTransaccion " & _
          "AND t.NPeriodo = i.Nperiodo AND t.NumeroMovimiento = i.NumeroMovimiento "
    If EnDolares Then Sql = Sql & "INNER JOIN Tasas ta ON t.FechaTransaccion = ta.FechaTasas "
    Sql = Sql & "WHERE t.FechaTransaccion < '" & SqlFecha(DateAdd("d", 1, Int(FechaHasta))) & "' "
    If SoloPeriodo Then Sql = Sql & "AND t.FechaTransaccion >= " & Ini & " "
    Sql = Sql & "AND c.CodCuentas BETWEEN '" & SqlTexto(CodDesde) & "' AND '" & SqlTexto(CodHasta) & "' " & _
                "AND ISNULL(i.Ajuste, '') <> N'" & SqlTexto(Ajuste) & "' " & _
                "GROUP BY c.CodCuentas, c.DescripcionCuentas, c.TipoCuenta ORDER BY c.CodCuentas"

    Set rsSql = New ADODB.Recordset
    rsSql.CursorLocation = adUseClient
    rsSql.Open Sql, cn, adOpenStatic, adLockReadOnly
    Set rsSql.ActiveConnection = Nothing
    cn.Close

    Set rs = NuevoRecordset()
    Do While Not rsSql.EOF
        Tipo = rsSql.Fields("TipoCuenta").Value & ""
        D1 = Dbl(rsSql.Fields("D1").Value)
        C1 = Dbl(rsSql.Fields("C1").Value)
        D2 = Round(Dbl(rsSql.Fields("D2").Value), 2)
        C2 = Round(Dbl(rsSql.Fields("C2").Value), 2)
        Debe1 = 0
        Haber1 = 0
        If Not SoloPeriodo Then
            If EsDeudora(Tipo) Then Net = Round(D1 - C1, 2) Else Net = Round(C1 - D1, 2)
            If EsDeudora(Tipo) Then
                If Net < 0 Then Haber1 = Abs(Net) Else Debe1 = Net
            Else
                If Net < 0 Then Debe1 = Abs(Net) Else Haber1 = Net
            End If
        End If
        If IsNull(rsSql.Fields("DescripcionCuentas").Value) Then
            Desc = rsSql.Fields("CodCuentas").Value & ".NO TIENE DESCRIPCION????"
        Else
            Desc = rsSql.Fields("CodCuentas").Value & "." & rsSql.Fields("DescripcionCuentas").Value
        End If
        Orden = Orden + 1
        rs.AddNew
        rs.Fields("Descripcion").Value = Replace(Desc, "'", "")
        rs.Fields("CodCuentas").Value = rsSql.Fields("CodCuentas").Value & ""
        rs.Fields("TipoCuenta").Value = Tipo
        rs.Fields("Debe1").Value = Debe1
        rs.Fields("Haber1").Value = Haber1
        rs.Fields("Debe2").Value = D2
        rs.Fields("Haber2").Value = C2
        If (Debe1 + D2) > (Haber1 + C2) Then rs.Fields("Debe3").Value = Round((Debe1 + D2) - (Haber1 + C2), 2) Else rs.Fields("Debe3").Value = 0
        If (Haber1 + C2) > (Debe1 + D2) Then rs.Fields("Haber3").Value = Round((Haber1 + C2) - (Debe1 + D2), 2) Else rs.Fields("Haber3").Value = 0
        rs.Fields("TotalDebe").Value = 0
        rs.Fields("TotalHaber").Value = 0
        rs.Fields("Nivel").Value = 0
        rs.Fields("Orden").Value = Orden
        rs.Update
        rsSql.MoveNext
    Loop
    rsSql.Close
    If rs.RecordCount > 0 Then rs.MoveFirst
    Set mRsBal = rs.Clone
    RptBalanzaActivo = True
    Set RptBalanzaCodigoGenerar = rs
    Exit Function

Falla:
    ErrNum = Err.Number: ErrDesc = Err.Description: ErrSrc = Err.Source
    On Error Resume Next
    cn.Close
    On Error GoTo 0
    Err.Raise ErrNum, ErrSrc, ErrDesc
End Function

'-----------------------------------------------------------------------------------------------------
' RptBalanzaGrupoGenerar  (BALANZA DE COMPROBACION - Por Grupo)
'   Arbol de Grupos (KeyGrupo entre GrupoDesde y GrupoHasta; vacio = todos): encabezado del grupo, sus
'   cuentas, los subgrupos y "Total <grupo>". Las columnas 1 y 2 de cada cuenta se acumulan en su grupo y en
'   todos los grupos superiores; la columna 3 se recalcula con el saldo neto de cada fila.
'   Como EliminaRegistroCero("Balanza"): se omiten las cuentas sin movimiento ni saldo final y los grupos
'   que se quedan sin cuentas. Los parametros de periodo, moneda y cierre son los de Por Codigo.
'-----------------------------------------------------------------------------------------------------
Public Function RptBalanzaGrupoGenerar(ByVal FechaDesde As Date, ByVal FechaHasta As Date, _
        ByVal EnDolares As Boolean, ByVal ExcluirCierre As Boolean, ByVal SoloPeriodo As Boolean, _
        ByVal GrupoDesde As String, ByVal GrupoHasta As String) As ADODB.Recordset
    Dim cn As ADODB.Connection, rsSql As ADODB.Recordset, rsG As ADODB.Recordset, Sql As String
    Dim Ini As String, Debito As String, Credito As String, CondPeriodo As String, Ajuste As String
    Dim ErrNum As Long, ErrDesc As String, ErrSrc As String
    Dim NG As Long, NA As Long, g As Long, a As Long, s As Long
    Dim gKey() As String, gSup() As String, gDesc() As String, gNivel() As Integer, gDatos() As Boolean
    Dim gAcc() As Double                         'gAcc(g, 0..3) = Debe1, Haber1, Debe2, Haber2
    Dim aCod() As String, aDesc() As String, aTipo() As String, aGrp() As Long
    Dim aVal() As Double                         'aVal(0..3, a) = Debe1, Haber1, Debe2, Haber2
    Dim aOk() As Boolean
    Dim dG As Object, D1 As Double, C1 As Double, D2 As Double, C2 As Double, Net As Double
    Dim Debe1 As Double, Haber1 As Double, Tipo As String, Desc As String, K As Long
    Dim rs As ADODB.Recordset, Orden As Long, V(0 To 3) As Double

    RptBalanzaLiberar
    If Len(GrupoHasta) = 0 Then GrupoHasta = "zzzzzzzzzzzzzzzz"
    If EnDolares Then Ajuste = "C" & ChrW(243) & "rdobas" Else Ajuste = "D" & ChrW(243) & "lares"
    Ini = "'" & SqlFecha(FechaDesde) & "'"
    If EnDolares Then
        Debito = "t.Debito * t.TCambio / ta.MontoCordobas"
        Credito = "t.Credito * t.TCambio / ta.MontoCordobas"
    Else
        Debito = "t.Debito * t.TCambio"
        Credito = "t.Credito * t.TCambio"
    End If
    CondPeriodo = "t.FechaTransaccion >= " & Ini
    If ExcluirCierre Then CondPeriodo = CondPeriodo & " AND ISNULL(i.Fuente, '') <> 'Cierre'"

    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open Conexion
    On Error GoTo Falla
    If EnDolares Then ValidaTasas cn, FechaDesde, FechaHasta, SoloPeriodo

    '--- grupos
    Set rsG = cn.Execute("SELECT KeyGrupo, KeyGrupoSuperior, DescripcionGrupo FROM Grupos WHERE KeyGrupo BETWEEN '" & _
                         SqlTexto(GrupoDesde) & "' AND '" & SqlTexto(GrupoHasta) & "' ORDER BY KeyGrupo")
    Set dG = CreateObject("Scripting.Dictionary")
    dG.CompareMode = vbTextCompare
    Do While Not rsG.EOF
        ReDim Preserve gKey(0 To NG): ReDim Preserve gSup(0 To NG): ReDim Preserve gDesc(0 To NG)
        ReDim Preserve gNivel(0 To NG): ReDim Preserve gDatos(0 To NG)
        gKey(NG) = rsG.Fields("KeyGrupo").Value & ""
        gSup(NG) = rsG.Fields("KeyGrupoSuperior").Value & ""
        gDesc(NG) = rsG.Fields("DescripcionGrupo").Value & ""
        gNivel(NG) = NivelGrupo(gKey(NG))
        dG(gKey(NG)) = NG
        NG = NG + 1
        rsG.MoveNext
    Loop
    rsG.Close
    If NG = 0 Then Err.Raise vbObjectError + 2002, "BalanzaMemoria", "No hay grupos en el rango seleccionado."
    ReDim gAcc(0 To NG - 1, 0 To 3)

    '--- saldos por cuenta
    Sql = "SELECT c.CodCuentas, c.DescripcionCuentas, c.TipoCuenta, c.KeyGrupo, " & _
          "SUM(CASE WHEN t.FechaTransaccion < " & Ini & " THEN " & Debito & " END) AS D1, " & _
          "SUM(CASE WHEN t.FechaTransaccion < " & Ini & " THEN " & Credito & " END) AS C1, " & _
          "SUM(CASE WHEN " & CondPeriodo & " THEN " & Debito & " END) AS D2, " & _
          "SUM(CASE WHEN " & CondPeriodo & " THEN " & Credito & " END) AS C2 " & _
          "FROM Cuentas c INNER JOIN Transacciones t ON c.CodCuentas = t.CodCuentas " & _
          "INNER JOIN IndiceTransaccion i ON t.FechaTransaccion = i.FechaTransaccion " & _
          "AND t.NPeriodo = i.Nperiodo AND t.NumeroMovimiento = i.NumeroMovimiento "
    If EnDolares Then Sql = Sql & "INNER JOIN Tasas ta ON t.FechaTransaccion = ta.FechaTasas "
    Sql = Sql & "WHERE t.FechaTransaccion < '" & SqlFecha(DateAdd("d", 1, Int(FechaHasta))) & "' "
    If SoloPeriodo Then Sql = Sql & "AND t.FechaTransaccion >= " & Ini & " "
    Sql = Sql & "AND c.KeyGrupo BETWEEN '" & SqlTexto(GrupoDesde) & "' AND '" & SqlTexto(GrupoHasta) & "' " & _
                "AND ISNULL(i.Ajuste, '') <> N'" & SqlTexto(Ajuste) & "' " & _
                "GROUP BY c.CodCuentas, c.DescripcionCuentas, c.TipoCuenta, c.KeyGrupo ORDER BY c.KeyGrupo, c.CodCuentas"
    Set rsSql = New ADODB.Recordset
    rsSql.CursorLocation = adUseClient
    rsSql.Open Sql, cn, adOpenStatic, adLockReadOnly
    Set rsSql.ActiveConnection = Nothing
    cn.Close

    Do While Not rsSql.EOF
        Tipo = rsSql.Fields("TipoCuenta").Value & ""
        D1 = Dbl(rsSql.Fields("D1").Value)
        C1 = Dbl(rsSql.Fields("C1").Value)
        D2 = Round(Dbl(rsSql.Fields("D2").Value), 2)
        C2 = Round(Dbl(rsSql.Fields("C2").Value), 2)
        Debe1 = 0
        Haber1 = 0
        If Not SoloPeriodo Then
            If EsDeudora(Tipo) Then Net = Round(D1 - C1, 2) Else Net = Round(C1 - D1, 2)
            If EsDeudora(Tipo) Then
                If Net < 0 Then Haber1 = Abs(Net) Else Debe1 = Net
            Else
                If Net < 0 Then Debe1 = Abs(Net) Else Haber1 = Net
            End If
        End If
        'Se omite la cuenta sin movimiento ni saldo final (Debe2 + Haber2 + Debe3 + Haber3 = 0)
        If Abs(D2) + Abs(C2) + Abs(Round((Debe1 + D2) - (Haber1 + C2), 2)) > 0 And dG.Exists(rsSql.Fields("KeyGrupo").Value & "") Then
            ReDim Preserve aCod(0 To NA): ReDim Preserve aDesc(0 To NA): ReDim Preserve aTipo(0 To NA)
            ReDim Preserve aGrp(0 To NA)
            If IsNull(rsSql.Fields("DescripcionCuentas").Value) Then
                Desc = rsSql.Fields("CodCuentas").Value & ".NO TIENE DESCRIPCION????"
            Else
                Desc = rsSql.Fields("CodCuentas").Value & "." & rsSql.Fields("DescripcionCuentas").Value
            End If
            aCod(NA) = rsSql.Fields("CodCuentas").Value & ""
            aDesc(NA) = Replace(Desc, "'", "")
            aTipo(NA) = Tipo
            aGrp(NA) = dG(rsSql.Fields("KeyGrupo").Value & "")
            ReDim Preserve aVal(0 To 3, 0 To NA)
            aVal(0, NA) = Debe1: aVal(1, NA) = Haber1: aVal(2, NA) = D2: aVal(3, NA) = C2
            'acumula en el grupo y en todos sus superiores
            g = aGrp(NA)
            Do
                gDatos(g) = True
                For K = 0 To 3
                    gAcc(g, K) = gAcc(g, K) + aVal(K, NA)
                Next
                If Len(gSup(g)) > 0 Then
                    If dG.Exists(gSup(g)) Then g = dG(gSup(g)) Else Exit Do
                Else
                    Exit Do
                End If
            Loop
            NA = NA + 1
        End If
        rsSql.MoveNext
    Loop
    rsSql.Close

    '--- filas
    Set rs = NuevoRecordset()
    For g = 0 To NG - 1
        If Len(gSup(g)) = 0 Then
            EmiteGrupoBalanza rs, g, NG, NA, gKey, gSup, gDesc, gNivel, gDatos, gAcc, aCod, aDesc, aTipo, aGrp, aVal, Orden
        ElseIf Not dG.Exists(gSup(g)) Then
            EmiteGrupoBalanza rs, g, NG, NA, gKey, gSup, gDesc, gNivel, gDatos, gAcc, aCod, aDesc, aTipo, aGrp, aVal, Orden
        End If
    Next
    If rs.RecordCount > 0 Then rs.MoveFirst
    Set mRsBal = rs.Clone
    RptBalanzaActivo = True
    Set RptBalanzaGrupoGenerar = rs
    Exit Function

Falla:
    ErrNum = Err.Number: ErrDesc = Err.Description: ErrSrc = Err.Source
    On Error Resume Next
    cn.Close
    On Error GoTo 0
    Err.Raise ErrNum, ErrSrc, ErrDesc
End Function

Private Function NivelGrupo(ByVal Key As String) As Integer
    Dim L As Long
    L = Len(Key)
    If L <= 1 Then
        NivelGrupo = 1
    ElseIf L = 5 Then
        NivelGrupo = 2
    Else
        NivelGrupo = CInt((L - 5) / 2) + 2
    End If
End Function

' Agrega una fila de la balanza por grupo. Columna 3 = saldo neto de las columnas 1 y 2.
Private Sub FilaBalanza(rs As ADODB.Recordset, ByVal Desc As String, ByVal KeyGrupo As String, ByVal KeySup As String, _
        ByVal KeyCuenta As String, ByVal Nivel As Integer, ByVal Cod As String, ByVal Tipo As String, _
        ByVal Debe1 As Double, ByVal Haber1 As Double, ByVal Debe2 As Double, ByVal Haber2 As Double, _
        ByVal Orden As Long)
    rs.AddNew
    rs.Fields("Descripcion").Value = Desc
    rs.Fields("Debe1").Value = Debe1
    rs.Fields("Haber1").Value = Haber1
    rs.Fields("Debe2").Value = Debe2
    rs.Fields("Haber2").Value = Haber2
    If (Debe1 + Debe2) > (Haber1 + Haber2) Then rs.Fields("Debe3").Value = Round((Debe1 + Debe2) - (Haber1 + Haber2), 2) Else rs.Fields("Debe3").Value = 0
    If (Haber1 + Haber2) > (Debe1 + Debe2) Then rs.Fields("Haber3").Value = Round((Haber1 + Haber2) - (Debe1 + Debe2), 2) Else rs.Fields("Haber3").Value = 0
    rs.Fields("TotalDebe").Value = 0
    rs.Fields("TotalHaber").Value = 0
    rs.Fields("KeyGrupo").Value = KeyGrupo
    rs.Fields("KeyGrupoSuperior").Value = KeySup
    rs.Fields("KeyGrupoCuenta").Value = KeyCuenta
    rs.Fields("Nivel").Value = Nivel
    rs.Fields("Orden").Value = Orden
    rs.Fields("CodCuentas").Value = Cod
    rs.Fields("TipoCuenta").Value = Tipo
    rs.Update
End Sub

Private Sub EmiteGrupoBalanza(rs As ADODB.Recordset, ByVal g As Long, ByVal NG As Long, ByVal NA As Long, _
        gKey() As String, gSup() As String, gDesc() As String, gNivel() As Integer, gDatos() As Boolean, _
        gAcc() As Double, aCod() As String, aDesc() As String, aTipo() As String, aGrp() As Long, _
        aVal() As Double, Orden As Long)
    Dim a As Long, c As Long
    If Not gDatos(g) Then Exit Sub
    Orden = Orden + 1
    FilaBalanza rs, Space$(gNivel(g) - 1) & gDesc(g), gKey(g), gSup(g), "", gNivel(g), "", "", 0, 0, 0, 0, Orden
    For a = 0 To NA - 1
        If aGrp(a) = g Then
            Orden = Orden + 1
            FilaBalanza rs, Space$(gNivel(g)) & aDesc(a), aCod(a), "", gKey(g), gNivel(g) + 1, aCod(a), aTipo(a), _
                        aVal(0, a), aVal(1, a), aVal(2, a), aVal(3, a), Orden
        End If
    Next
    For c = 0 To NG - 1
        If c <> g Then
            If StrComp(gSup(c), gKey(g), vbTextCompare) = 0 Then
                EmiteGrupoBalanza rs, c, NG, NA, gKey, gSup, gDesc, gNivel, gDatos, gAcc, aCod, aDesc, aTipo, aGrp, aVal, Orden
            End If
        End If
    Next
    Orden = Orden + 1
    FilaBalanza rs, Space$(gNivel(g) - 1) & "Total " & gDesc(g), gKey(g), gSup(g), "", gNivel(g), "", "", _
                gAcc(g, 0), gAcc(g, 1), gAcc(g, 2), gAcc(g, 3), Orden
End Sub

' Suma de una columna (Debe1..Haber3) de las filas Nivel = 1 con "Total" en la descripcion: lo mismo que
' ArepBalanza pide a Reportes en GroupFooter1_Format.
Public Function RptBalanzaSuma(ByVal Campo As String) As Double
    Dim rs As ADODB.Recordset, Suma As Double
    If mRsBal Is Nothing Then Exit Function
    Set rs = mRsBal.Clone
    rs.Filter = "Nivel = 1 AND Descripcion LIKE '*Total*'"
    Do While Not rs.EOF
        Suma = Suma + Dbl(rs.Fields(Campo).Value)
        rs.MoveNext
    Loop
    rs.Close
    RptBalanzaSuma = Suma
End Function

Public Sub RptBalanzaLiberar()
    RptBalanzaActivo = False
    Set mRsBal = Nothing
End Sub
