Attribute VB_Name = "ResultadoDpto"
Option Explicit
'=====================================================================================================
' ESTADO DE RESULTADO POR DEPARTAMENTO - VERSION EN MEMORIA
'
' Funciones paralelas a CreaEstructuraDpto / SaldoReportesDpto (Funciones.bas). NO usan la tabla
' Reportes: una sola consulta SQL trae los saldos por Cuenta + Departamento y todo lo demas
' (estructura de grupos, totales por departamento, totales por grupo, filtro de nivel, ceros)
' se arma en memoria y se devuelve como un Recordset desconectado con las mismas columnas que
' usa ArepResultado (Descripcion, Debe1, Haber1, KeyGrupo, Nivel, Orden ...).
'
' Funciones publicas (reutilizables por otros reportes):
'   RptSaldosPorCuentaDpto          Saldos (Debito/Credito) por cuenta y departamento, 1 consulta.
'   RptResultadoDptoGenerar         Recordset completo del Estado de Resultado por Dpto.
'   RptVerificarCuadreResultadoDpto Prueba: consolidado vs suma departamento por departamento.
'
' Reglas de negocio (iguales al reporte original):
'   - Ingresos - Ventas  : Credito - Debito.   Costos y Gastos: Debito - Credito.
'   - Moneda Cordobas    : Debito * TCambio.   Dolares: Debito * TCambio / Tasas.MontoCordobas.
'   - Departamento       : Transacciones.VoucherNo. Nulo, vacio o "-" = sin departamento ("-").
'=====================================================================================================

Private Const EPS As Double = 0.000005
Private Const TIPO_INGRESOS As String = "Ingresos - Ventas"
Private Const TIPOS_RESULTADO As String = "'Ingresos - Ventas','Costos','Gastos'"

Private Type TGrupo
    Key As String
    Superior As String
    Descripcion As String
    Nivel As Integer
    Total As Double
    Datos As Boolean
    Visitado As Boolean
End Type

Private mGrp() As TGrupo
Private mNGrp As Long
Private mSal As Variant          'GetRows: (0 Cod, 1 Desc, 2 KeyGrupo, 3 Tipo, 4 Dpto, 5 Debito, 6 Credito)
Private mMonto() As Double       'monto con signo contable segun tipo de cuenta
Private mSalGrp() As Long        'indice de grupo de cada fila de mSal (-1 = cuenta fuera de la estructura)
Private mNSal As Long
Private mRs As ADODB.Recordset
Private mOrden As Long
Private mNivelMax As Integer
Private mOcultarCeros As Boolean
Private mDptoNombre As Object


'-----------------------------------------------------------------------------------------------------
' Utilidades SQL
'-----------------------------------------------------------------------------------------------------
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

' Expresion SQL del departamento de la transaccion.
' ParaRango = True  : "sin departamento" se compara como '00' (igual que el filtro original Desde/Hasta).
' ParaRango = False : "sin departamento" se devuelve como '-'.
Private Function SqlDpto(ByVal ParaRango As Boolean) As String
    Dim Vacio As String
    If ParaRango Then Vacio = "'00'" Else Vacio = "'-'"
    SqlDpto = "CASE WHEN t.VoucherNo IS NULL THEN " & Vacio & _
              " WHEN LTRIM(RTRIM(t.VoucherNo)) IN ('', '-') THEN " & Vacio & _
              " ELSE LTRIM(RTRIM(t.VoucherNo)) END"
End Function

' FROM ... WHERE comun a todas las consultas del reporte.
Private Function SqlDesdeDonde(ByVal FechaDesde As Date, ByVal FechaHasta As Date, _
        ByVal EnDolares As Boolean, ByVal ExcluirCierre As Boolean, _
        ByVal DptoDesde As String, ByVal DptoHasta As String, ByVal TiposCuenta As String) As String
    Dim Sql As String
    Sql = "FROM Cuentas c INNER JOIN Transacciones t ON c.CodCuentas = t.CodCuentas "
    If ExcluirCierre Then
        Sql = Sql & "INNER JOIN IndiceTransaccion i ON t.FechaTransaccion = i.FechaTransaccion " & _
                    "AND t.NPeriodo = i.Nperiodo AND t.NumeroMovimiento = i.NumeroMovimiento "
    End If
    If EnDolares Then
        Sql = Sql & "INNER JOIN Tasas ta ON t.FechaTransaccion = ta.FechaTasas "
    End If
    Sql = Sql & "WHERE t.FechaTransaccion >= '" & SqlFecha(FechaDesde) & "' " & _
                "AND t.FechaTransaccion < '" & SqlFecha(DateAdd("d", 1, Int(FechaHasta))) & "' " & _
                "AND c.TipoCuenta IN (" & TiposCuenta & ") "
    If ExcluirCierre Then
        Sql = Sql & "AND ISNULL(i.Fuente, '') <> 'Cierre' "
    End If
    If Len(DptoDesde) > 0 Or Len(DptoHasta) > 0 Then
        If Len(DptoHasta) = 0 Then DptoHasta = "zzzzzzzzzz"
        Sql = Sql & "AND (" & SqlDpto(True) & ") BETWEEN '" & SqlTexto(DptoDesde) & "' AND '" & SqlTexto(DptoHasta) & "' "
    End If
    SqlDesdeDonde = Sql
End Function

Private Function SqlImporte(ByVal Campo As String, ByVal EnDolares As Boolean) As String
    If EnDolares Then
        SqlImporte = "t." & Campo & " * t.TCambio / ta.MontoCordobas"
    Else
        SqlImporte = "t." & Campo & " * t.TCambio"
    End If
End Function

' Falla si hay movimientos del periodo sin tasa de cambio (el reporte original abortaba igual).
Private Sub ValidaTasas(cn As ADODB.Connection, ByVal FechaDesde As Date, ByVal FechaHasta As Date)
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("SELECT COUNT(*) FROM Transacciones t INNER JOIN Cuentas c ON c.CodCuentas = t.CodCuentas " & _
                        "LEFT JOIN Tasas ta ON t.FechaTransaccion = ta.FechaTasas " & _
                        "WHERE c.TipoCuenta IN (" & TIPOS_RESULTADO & ") AND t.FechaTransaccion >= '" & SqlFecha(FechaDesde) & "' " & _
                        "AND t.FechaTransaccion < '" & SqlFecha(DateAdd("d", 1, Int(FechaHasta))) & "' " & _
                        "AND (ta.MontoCordobas IS NULL OR ta.MontoCordobas = 0)")
    If CLng(rs.Fields(0).Value) > 0 Then
        rs.Close
        Err.Raise vbObjectError + 1001, "ResultadoDpto", _
            "Existen movimientos en el periodo sin Tasa de Cambio (o con tasa cero)." & vbLf & _
            "El reporte en Dolares no puede continuar."
    End If
    rs.Close
End Sub


'-----------------------------------------------------------------------------------------------------
' RptSaldosPorCuentaDpto
' Una sola consulta: Debito y Credito (en la moneda pedida) por Cuenta + Departamento.
' Recordset desconectado, ordenado por KeyGrupo, Departamento, Cuenta.
' Columnas: CodCuentas, DescripcionCuentas, KeyGrupo, TipoCuenta, CodDepartamento, Debito, Credito
'-----------------------------------------------------------------------------------------------------
Public Function RptSaldosPorCuentaDpto(ByVal FechaDesde As Date, ByVal FechaHasta As Date, _
        ByVal EnDolares As Boolean, ByVal ExcluirCierre As Boolean, _
        Optional ByVal DptoDesde As String = "", Optional ByVal DptoHasta As String = "", _
        Optional ByVal TiposCuenta As String = "'Ingresos - Ventas','Costos','Gastos'") As ADODB.Recordset
    Dim cn As ADODB.Connection, rs As ADODB.Recordset, Sql As String
    Dim ErrNum As Long, ErrDesc As String, ErrSrc As String

    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open Conexion
    On Error GoTo Falla

    If EnDolares Then ValidaTasas cn, FechaDesde, FechaHasta

    Sql = "SELECT c.CodCuentas, c.DescripcionCuentas, c.KeyGrupo, c.TipoCuenta, " & SqlDpto(False) & " AS CodDepartamento, " & _
          "SUM(" & SqlImporte("Debito", EnDolares) & ") AS Debito, " & _
          "SUM(" & SqlImporte("Credito", EnDolares) & ") AS Credito " & _
          SqlDesdeDonde(FechaDesde, FechaHasta, EnDolares, ExcluirCierre, DptoDesde, DptoHasta, TiposCuenta) & _
          "GROUP BY c.CodCuentas, c.DescripcionCuentas, c.KeyGrupo, c.TipoCuenta, " & SqlDpto(False) & " " & _
          "ORDER BY c.KeyGrupo, " & SqlDpto(False) & ", c.CodCuentas"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Sql, cn, adOpenStatic, adLockReadOnly
    Set rs.ActiveConnection = Nothing
    cn.Close
    Set RptSaldosPorCuentaDpto = rs
    Exit Function

Falla:
    ErrNum = Err.Number: ErrDesc = Err.Description: ErrSrc = Err.Source
    On Error Resume Next
    cn.Close
    On Error GoTo 0
    Err.Raise ErrNum, ErrSrc, ErrDesc
End Function


'-----------------------------------------------------------------------------------------------------
' Estructura de grupos y nombres de departamentos (consultas pequenas, una vez por reporte)
'-----------------------------------------------------------------------------------------------------
Private Sub CargarGrupos(ByVal Prefijos As String)
    Dim cn As ADODB.Connection, rs As ADODB.Recordset, Sql As String, P As Variant, Filtro As String

    For Each P In VBA.Split(Prefijos, ",")
        If Len(Trim$(P)) > 0 Then
            If Len(Filtro) > 0 Then Filtro = Filtro & " OR "
            Filtro = Filtro & "KeyGrupo LIKE '" & SqlTexto(Trim$(P)) & "%'"
        End If
    Next
    Sql = "SELECT KeyGrupo, KeyGrupoSuperior, DescripcionGrupo FROM Grupos WHERE (" & Filtro & ") ORDER BY KeyGrupo"

    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open Conexion
    Set rs = cn.Execute(Sql)

    mNGrp = 0
    ReDim mGrp(0 To 0)
    Do While Not rs.EOF
        ReDim Preserve mGrp(0 To mNGrp)
        With mGrp(mNGrp)
            .Key = rs.Fields("KeyGrupo").Value & ""
            .Superior = rs.Fields("KeyGrupoSuperior").Value & ""
            .Descripcion = rs.Fields("DescripcionGrupo").Value & ""
            .Nivel = NivelDeGrupo(.Key)
            .Total = 0
            .Datos = False
            .Visitado = False
        End With
        mNGrp = mNGrp + 1
        rs.MoveNext
    Loop
    rs.Close

    Set mDptoNombre = CreateObject("Scripting.Dictionary")
    mDptoNombre.CompareMode = vbTextCompare
    Set rs = cn.Execute("SELECT CodGrupo, DescripcionGrupo FROM GrupoCuentas")
    Do While Not rs.EOF
        mDptoNombre(Trim$(rs.Fields("CodGrupo").Value & "")) = rs.Fields("DescripcionGrupo").Value & ""
        rs.MoveNext
    Loop
    rs.Close
    cn.Close
End Sub

' Mismo calculo de nivel que usa el reporte original (largo de KeyGrupo).
Private Function NivelDeGrupo(ByVal Key As String) As Integer
    Dim L As Long
    L = Len(Key)
    If L <= 1 Then
        NivelDeGrupo = 1
    ElseIf L = 5 Then
        NivelDeGrupo = 2
    Else
        NivelDeGrupo = CInt((L - 5) / 2) + 2
    End If
End Function

Private Function NombreDpto(ByVal Cod As String) As String
    If Cod = "-" Then
        NombreDpto = "Sin Departamento"
    ElseIf mDptoNombre.Exists(Cod) Then
        NombreDpto = mDptoNombre(Cod)
    Else
        NombreDpto = "Departamento " & Cod & " (no existe en catalogo)"
    End If
End Function


'-----------------------------------------------------------------------------------------------------
' RptResultadoDptoGenerar
'   NivelMax      0 = todos los niveles (igual que CmbNivel)
'   OcultarCeros  True = no muestra cuentas, departamentos ni grupos sin movimiento
'   Advertencias  (salida) cuentas con saldo que no estan en la estructura de grupos
' Devuelve un Recordset desconectado listo para ArepResultado.
'-----------------------------------------------------------------------------------------------------
Public Function RptResultadoDptoGenerar(ByVal FechaDesde As Date, ByVal FechaHasta As Date, _
        ByVal EnDolares As Boolean, ByVal ExcluirCierre As Boolean, _
        ByVal DptoDesde As String, ByVal DptoHasta As String, _
        ByVal NivelMax As Integer, ByVal OcultarCeros As Boolean, _
        Optional ByRef Advertencias As String = "", _
        Optional ByVal Prefijos As String = "D,G,O") As ADODB.Recordset
    Dim rsSal As ADODB.Recordset, dKeys As Object
    Dim i As Long, g As Long, Clave As String
    Dim IngresosTot As Double, GastosTot As Double, SinGrupo As Double

    Advertencias = ""
    mNivelMax = NivelMax
    mOcultarCeros = OcultarCeros
    mOrden = 0

    Set rsSal = RptSaldosPorCuentaDpto(FechaDesde, FechaHasta, EnDolares, ExcluirCierre, DptoDesde, DptoHasta)
    CargarGrupos Prefijos

    Set dKeys = CreateObject("Scripting.Dictionary")
    dKeys.CompareMode = vbTextCompare
    For g = 0 To mNGrp - 1
        dKeys(mGrp(g).Key) = g
    Next

    '--- saldos a memoria, con signo contable y acumulado en el grupo propio de cada cuenta
    mNSal = 0
    If Not (rsSal.BOF And rsSal.EOF) Then
        mNSal = rsSal.RecordCount
        mSal = rsSal.GetRows
    End If
    rsSal.Close
    If mNSal > 0 Then
        ReDim mMonto(0 To mNSal - 1)
        ReDim mSalGrp(0 To mNSal - 1)
    Else
        ReDim mMonto(0 To 0)
        ReDim mSalGrp(0 To 0)
    End If

    For i = 0 To mNSal - 1
        If (mSal(3, i) & "") = TIPO_INGRESOS Then
            mMonto(i) = Dbl(mSal(6, i)) - Dbl(mSal(5, i))
            IngresosTot = IngresosTot + mMonto(i)
        Else
            mMonto(i) = Dbl(mSal(5, i)) - Dbl(mSal(6, i))
            GastosTot = GastosTot + mMonto(i)
        End If
        Clave = mSal(2, i) & ""
        If dKeys.Exists(Clave) Then
            g = dKeys(Clave)
            mSalGrp(i) = g
            mGrp(g).Total = mGrp(g).Total + mMonto(i)
            If Abs(mMonto(i)) > EPS Then mGrp(g).Datos = True
        Else
            mSalGrp(i) = -1
            SinGrupo = SinGrupo + mMonto(i)
            Advertencias = Advertencias & "Cuenta " & mSal(0, i) & " (grupo '" & Clave & "', dpto " & mSal(4, i) & _
                           ") tiene saldo " & Format$(mMonto(i), "#,##0.00") & " y no esta en la estructura." & vbLf
        End If
    Next

    '--- acumula hijos en padres
    For g = 0 To mNGrp - 1
        If EsRaiz(g, dKeys) Then CalcularGrupo g
    Next
    For g = 0 To mNGrp - 1
        mGrp(g).Visitado = False
    Next

    '--- arma el recordset
    Set mRs = NuevoRecordset()
    For g = 0 To mNGrp - 1
        If EsRaiz(g, dKeys) Then EmitirGrupo g
    Next

    AgregaFila " Total Costos y Gastos", "CG", "", "", 1, 0, GastosTot, True, "", ""
    AgregaFila " Resultado Periodo", "RP", "", "", 1, 0, IngresosTot - GastosTot, True, "", ""

    If SinGrupo <> 0 Then
        Advertencias = Advertencias & "El Resultado Periodo incluye " & Format$(SinGrupo, "#,##0.00") & _
                       " de cuentas que no aparecen en la estructura." & vbLf
    End If

    If mRs.RecordCount > 0 Then mRs.MoveFirst
    Set RptResultadoDptoGenerar = mRs
    Set mRs = Nothing
    mSal = Empty
End Function

Private Function EsRaiz(ByVal g As Long, dKeys As Object) As Boolean
    EsRaiz = (Len(mGrp(g).Superior) = 0) Or Not dKeys.Exists(mGrp(g).Superior)
End Function

' Suma recursivamente los totales de los hijos en el padre.
Private Sub CalcularGrupo(ByVal g As Long)
    Dim c As Long
    If mGrp(g).Visitado Then Exit Sub
    mGrp(g).Visitado = True
    For c = 0 To mNGrp - 1
        If c <> g Then
            If mGrp(c).Superior = mGrp(g).Key And Not mGrp(c).Visitado Then
                CalcularGrupo c
                mGrp(g).Total = mGrp(g).Total + mGrp(c).Total
                If mGrp(c).Datos Then mGrp(g).Datos = True
            End If
        End If
    Next
End Sub

Private Sub EmitirGrupo(ByVal g As Long)
    Dim c As Long
    If mGrp(g).Visitado Then Exit Sub
    mGrp(g).Visitado = True
    If mOcultarCeros And Not mGrp(g).Datos Then Exit Sub

    AgregaFila Space$(mGrp(g).Nivel - 1) & mGrp(g).Descripcion, mGrp(g).Key, mGrp(g).Superior, "", _
               mGrp(g).Nivel, 0, 0, False, "", ""

    EmitirCuentasGrupo g

    For c = 0 To mNGrp - 1
        If c <> g Then
            If mGrp(c).Superior = mGrp(g).Key Then EmitirGrupo c
        End If
    Next

    AgregaFila Space$(mGrp(g).Nivel - 1) & "Total " & mGrp(g).Descripcion, mGrp(g).Key, mGrp(g).Superior, "", _
               mGrp(g).Nivel, 0, mGrp(g).Total, True, "", ""
End Sub

' Cuentas del grupo agrupadas por departamento: encabezado, cuentas, total del departamento.
Private Sub EmitirCuentasGrupo(ByVal g As Long)
    Dim i As Long, Dpto As String, DptoActual As String, Abierto As Boolean, TotDpto As Double
    Dim Nivel As Integer, Key As String, Sup As String

    Nivel = mGrp(g).Nivel
    Key = mGrp(g).Key
    Sup = mGrp(g).Superior

    For i = 0 To mNSal - 1
        If mSalGrp(i) = g Then
            If Not (mOcultarCeros And Abs(mMonto(i)) < EPS) Then
                Dpto = mSal(4, i) & ""
                If Abierto And Dpto <> DptoActual Then
                    AgregaFila Space$(Nivel) & "Total " & NombreDpto(DptoActual), Key, Sup, "", Nivel + 1, 0, TotDpto, True, DptoActual, ""
                    Abierto = False
                End If
                If Not Abierto Then
                    AgregaFila Space$(Nivel) & NombreDpto(Dpto), Key, Sup, "", Nivel + 1, 0, 0, False, Dpto, ""
                    DptoActual = Dpto
                    TotDpto = 0
                    Abierto = True
                End If
                AgregaFila Space$(Nivel) & mSal(0, i) & "." & (mSal(1, i) & ""), mSal(0, i) & "", "", Key, Nivel + 1, _
                           mMonto(i), 0, False, Dpto, mSal(0, i) & ""
                TotDpto = TotDpto + mMonto(i)
            End If
        End If
    Next
    If Abierto Then
        AgregaFila Space$(Nivel) & "Total " & NombreDpto(DptoActual), Key, Sup, "", Nivel + 1, 0, TotDpto, True, DptoActual, ""
    End If
End Sub


'-----------------------------------------------------------------------------------------------------
' Recordset en memoria
'-----------------------------------------------------------------------------------------------------
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
        .Append "KeyGrupo", adVarWChar, 50, adFldIsNullable
        .Append "KeyGrupoSuperior", adVarWChar, 50, adFldIsNullable
        .Append "KeyGrupoCuenta", adVarWChar, 50, adFldIsNullable
        .Append "Nivel", adInteger
        .Append "Orden", adInteger
        .Append "CodDepartamento", adVarWChar, 50, adFldIsNullable
        .Append "CodCuentas", adVarWChar, 50, adFldIsNullable
    End With
    rs.Open
    Set NuevoRecordset = rs
End Function

' Agrega una fila aplicando el filtro de nivel (misma regla que CmbNivel en el reporte original:
' se quita todo lo mayor al nivel y, en el nivel elegido, solo se dejan los "Total").
Private Sub AgregaFila(ByVal Descripcion As String, ByVal KeyGrupo As String, ByVal KeySuperior As String, _
        ByVal KeyCuenta As String, ByVal Nivel As Integer, ByVal Debe1 As Double, ByVal Haber1 As Double, _
        ByVal EsTotal As Boolean, ByVal CodDpto As String, ByVal CodCuenta As String)
    If mNivelMax > 0 Then
        If Nivel > mNivelMax Then Exit Sub
        If Nivel = mNivelMax And Not EsTotal Then Exit Sub
    End If
    mOrden = mOrden + 1
    mRs.AddNew
    mRs.Fields("Descripcion").Value = Descripcion
    mRs.Fields("Debe1").Value = Debe1
    mRs.Fields("Haber1").Value = Haber1
    mRs.Fields("Debe2").Value = 0
    mRs.Fields("Haber2").Value = 0
    mRs.Fields("Debe3").Value = 0
    mRs.Fields("Haber3").Value = 0
    mRs.Fields("KeyGrupo").Value = KeyGrupo
    mRs.Fields("KeyGrupoSuperior").Value = KeySuperior
    mRs.Fields("KeyGrupoCuenta").Value = KeyCuenta
    mRs.Fields("Nivel").Value = Nivel
    mRs.Fields("Orden").Value = mOrden
    mRs.Fields("CodDepartamento").Value = CodDpto
    mRs.Fields("CodCuentas").Value = CodCuenta
    mRs.Update
End Sub


'-----------------------------------------------------------------------------------------------------
' RptVerificarCuadreResultadoDpto
' Prueba de cuadre. Uso desde la ventana Inmediato (Ctrl+G) con la base abierta:
'   ?RptVerificarCuadreResultadoDpto(#1/1/2024#, #1/31/2024#, False, False)
' Compara el Resultado del Periodo:
'   (a) consulta directa sin agrupar por departamento
'   (b) suma de todos los departamentos juntos
'   (c) cada departamento por separado con el filtro Desde/Hasta
'-----------------------------------------------------------------------------------------------------
Public Function RptVerificarCuadreResultadoDpto(ByVal FechaDesde As Date, ByVal FechaHasta As Date, _
        ByVal EnDolares As Boolean, ByVal ExcluirCierre As Boolean) As String
    Dim cn As ADODB.Connection, rs As ADODB.Recordset, Sql As String
    Dim Directo As Double, Suma As Double, SumaDptos As Double, Texto As String
    Dim d As Object, K As Variant, Monto As Double, Net As Double

    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open Conexion
    If EnDolares Then ValidaTasas cn, FechaDesde, FechaHasta
    Sql = "SELECT SUM(CASE WHEN c.TipoCuenta = '" & TIPO_INGRESOS & "' THEN (" & SqlImporte("Credito", EnDolares) & ") - (" & SqlImporte("Debito", EnDolares) & ") " & _
          "ELSE (" & SqlImporte("Debito", EnDolares) & ") - (" & SqlImporte("Credito", EnDolares) & ") END * " & _
          "CASE WHEN c.TipoCuenta = '" & TIPO_INGRESOS & "' THEN 1 ELSE -1 END) " & _
          SqlDesdeDonde(FechaDesde, FechaHasta, EnDolares, ExcluirCierre, "", "", TIPOS_RESULTADO)
    Set rs = cn.Execute(Sql)
    If Not IsNull(rs.Fields(0).Value) Then Directo = CDbl(rs.Fields(0).Value)
    rs.Close
    cn.Close

    Set d = CreateObject("Scripting.Dictionary")
    Set rs = RptSaldosPorCuentaDpto(FechaDesde, FechaHasta, EnDolares, ExcluirCierre)
    Do While Not rs.EOF
        If (rs.Fields("TipoCuenta").Value & "") = TIPO_INGRESOS Then
            Net = Dbl(rs.Fields("Credito").Value) - Dbl(rs.Fields("Debito").Value)
        Else
            Net = -(Dbl(rs.Fields("Debito").Value) - Dbl(rs.Fields("Credito").Value))
        End If
        Suma = Suma + Net
        d(rs.Fields("CodDepartamento").Value & "") = d(rs.Fields("CodDepartamento").Value & "") + Net
        rs.MoveNext
    Loop
    rs.Close

    Texto = "Resultado consulta directa (sin dpto): " & Format$(Directo, "#,##0.00") & vbLf & _
            "Resultado agrupado por dpto         : " & Format$(Suma, "#,##0.00") & vbLf

    For Each K In d.Keys
        Set rs = RptSaldosPorCuentaDpto(FechaDesde, FechaHasta, EnDolares, ExcluirCierre, IIf(K = "-", "00", CStr(K)), IIf(K = "-", "00", CStr(K)))
        Monto = 0
        Do While Not rs.EOF
            If (rs.Fields("TipoCuenta").Value & "") = TIPO_INGRESOS Then
                Monto = Monto + Dbl(rs.Fields("Credito").Value) - Dbl(rs.Fields("Debito").Value)
            Else
                Monto = Monto - (Dbl(rs.Fields("Debito").Value) - Dbl(rs.Fields("Credito").Value))
            End If
            rs.MoveNext
        Loop
        rs.Close
        SumaDptos = SumaDptos + Monto
        Texto = Texto & "  Dpto " & K & ": " & Format$(d(K), "#,##0.00") & " | filtrado: " & Format$(Monto, "#,##0.00")
        If Abs(Monto - d(K)) > 0.01 Then Texto = Texto & "   <-- DIFERENCIA"
        Texto = Texto & vbLf
    Next

    Texto = Texto & "Suma de departamentos filtrados  : " & Format$(SumaDptos, "#,##0.00") & vbLf
    If Abs(Directo - Suma) > 0.01 Or Abs(Directo - SumaDptos) > 0.01 Then
        Texto = Texto & "*** NO CUADRA ***"
    Else
        Texto = Texto & "OK: cuadra."
    End If
    RptVerificarCuadreResultadoDpto = Texto
End Function
