Attribute VB_Name = "ResultadoDpto"
Option Explicit
'=====================================================================================================
' ESTADOS DE RESULTADO - VERSION EN MEMORIA (por Departamento, normal, acumulado e historico)
'
' Funciones paralelas a CreaEstructuraDpto / SaldoReportesDpto (Funciones.bas). NO usan la tabla
' Reportes: una sola consulta SQL trae los saldos por Cuenta + Departamento y todo lo demas
' (estructura de grupos, totales por departamento, totales por grupo, filtro de nivel, ceros)
' se arma en memoria y se devuelve como un Recordset desconectado con las mismas columnas que
' usa ArepResultado (Descripcion, Debe1, Haber1, KeyGrupo, Nivel, Orden ...).
'
' Funciones publicas (reutilizables por otros reportes):
'   RptSaldosPorCuentaDpto          Saldos (Debito/Credito) por cuenta y departamento, 1 consulta.
'   RptResultadoDptoGenerar         Recordset completo del Estado de Resultado (con o sin Dpto, con o sin
'                                   columna de saldo anterior).
'   RptResultadoGenerar             Atajo sin Dpto (ESTADO DE RESULTADO / RESULTADO ACUMULADO / HISTORICO).
'   RptConvertirRecordset           Divide el reporte entre una tasa (equivale a ConvertirReporte, en memoria).
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
Private Const TIPOS_BALANCE As String = "'Otros Activos','Caja','Bancos','Cuentas x Cobrar','Inventario','Papeleria - Utiles','Activo Fijo','Otros Pasivos','Cuentas x Pagar','Pasivo','Capital'"

Private Type TGrupo
    Key As String
    Superior As String
    Descripcion As String
    Nivel As Integer
    Total As Double
    Total2 As Double
    Datos As Boolean
    Visitado As Boolean
End Type

Private mGrp() As TGrupo
Private mNGrp As Long
Private mSal As Variant          'GetRows: (0 Cod, 1 Desc, 2 KeyGrupo, 3 Tipo, 4 Dpto, 5 Debito, 6 Credito, 7 Debito2, 8 Credito2, 9 UbicacionReporte)
Private mMonto() As Double       'monto con signo contable segun tipo de cuenta
Private mSalGrp() As Long        'indice de grupo de cada fila de mSal (-1 = cuenta fuera de la estructura)
Private mNSal As Long
Private mRs As ADODB.Recordset
Private mOrden As Long
Private mNivelMax As Integer
Private mOcultarCeros As Boolean
Private mDptoNombre As Object
Private mMonto2() As Double      'monto del saldo anterior (< FechaDesde); solo con ConAnterior
Private mPorDpto As Boolean
Private mConAnterior As Boolean
Private mBalance As Boolean
Private mResul1 As Double        'Balance: resultado acumulado anterior (columna 1)
Private mResul2 As Double        'Balance: resultado del periodo (columna 2)
Private mUbicSaldo As Object     'Balance: Ubicacion -> saldo (col 1 + col 2), para SaldosPersonalizados
Private mUbicPeriodo As Object   'Balance: Ubicacion -> actividad del periodo (col 2)
Private mFilaUbic As String      'datos extra de la fila de cuenta que se esta agregando (AgregaFila los consume)
Private mFilaTipo As String
Private mFilaDescCta As String
Private mRsAnexos As ADODB.Recordset
Private mRP2 As Double           'Resultado Periodo: Haber2 - Debe2 de la fila RP
Private mRP3 As Double           'Resultado Periodo: Haber3 - Debe3 de la fila RP

'Verdadero mientras un Balance Resumen (NUEVO) esta imprimiendo: SaldosPersonalizados y ArepAnexosBalances
'leen de memoria en vez de la tabla Reportes. El formulario lo apaga con RptUbicacionesLiberar.
Public RptUbicActivo As Boolean


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
        ByVal DptoDesde As String, ByVal DptoHasta As String, ByVal TiposCuenta As String, _
        Optional ByVal SinLimiteInferior As Boolean = False, Optional ByVal AjusteExcluir As String = "") As String
    Dim Sql As String
    Sql = "FROM Cuentas c INNER JOIN Transacciones t ON c.CodCuentas = t.CodCuentas "
    If ExcluirCierre Or Len(AjusteExcluir) > 0 Then
        Sql = Sql & "INNER JOIN IndiceTransaccion i ON t.FechaTransaccion = i.FechaTransaccion " & _
                    "AND t.NPeriodo = i.Nperiodo AND t.NumeroMovimiento = i.NumeroMovimiento "
    End If
    If EnDolares Then
        Sql = Sql & "INNER JOIN Tasas ta ON t.FechaTransaccion = ta.FechaTasas "
    End If
    Sql = Sql & "WHERE "
    If Not SinLimiteInferior Then Sql = Sql & "t.FechaTransaccion >= '" & SqlFecha(FechaDesde) & "' AND "
    Sql = Sql & "t.FechaTransaccion < '" & SqlFecha(DateAdd("d", 1, Int(FechaHasta))) & "' " & _
                "AND c.TipoCuenta IN (" & TiposCuenta & ") "
    If ExcluirCierre Then
        Sql = Sql & "AND ISNULL(i.Fuente, '') <> 'Cierre' "
    End If
    'Asientos de ajuste de la otra moneda (IndiceTransaccion.Ajuste), como hace el reporte acumulado original
    If Len(AjusteExcluir) > 0 Then
        Sql = Sql & "AND ISNULL(i.Ajuste, '') <> N'" & SqlTexto(AjusteExcluir) & "' "
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
Private Sub ValidaTasas(cn As ADODB.Connection, ByVal FechaDesde As Date, ByVal FechaHasta As Date, Optional ByVal Tipos As String = "'Ingresos - Ventas','Costos','Gastos'")
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("SELECT COUNT(*) FROM Transacciones t INNER JOIN Cuentas c ON c.CodCuentas = t.CodCuentas " & _
                        "LEFT JOIN Tasas ta ON t.FechaTransaccion = ta.FechaTasas " & _
                        "WHERE c.TipoCuenta IN (" & Tipos & ") AND t.FechaTransaccion >= '" & SqlFecha(FechaDesde) & "' " & _
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
' Columnas: CodCuentas, DescripcionCuentas, KeyGrupo, TipoCuenta, CodDepartamento, Debito, Credito, Debito2, Credito2
'   PorDpto = False  : no agrupa por departamento (CodDepartamento = "-")
'   ConAnterior      : Debito/Credito = movimientos anteriores a FechaDesde (saldo anterior, columna 1) y
'                      Debito2/Credito2 = actividad FechaDesde..FechaHasta (columna 2), como SaldoReportesAcumulado.
'                      Sin ConAnterior: Debito/Credito = actividad del periodo y Debito2/Credito2 = 0.
'   AjusteExcluir    : si trae texto ("Dolares"/"Cordobas") excluye los asientos con ese IndiceTransaccion.Ajuste
'-----------------------------------------------------------------------------------------------------
Public Function RptSaldosPorCuentaDpto(ByVal FechaDesde As Date, ByVal FechaHasta As Date, _
        ByVal EnDolares As Boolean, ByVal ExcluirCierre As Boolean, _
        Optional ByVal DptoDesde As String = "", Optional ByVal DptoHasta As String = "", _
        Optional ByVal TiposCuenta As String = "'Ingresos - Ventas','Costos','Gastos'", _
        Optional ByVal PorDpto As Boolean = True, Optional ByVal ConAnterior As Boolean = False, _
        Optional ByVal AjusteExcluir As String = "") As ADODB.Recordset
    Dim cn As ADODB.Connection, rs As ADODB.Recordset, Sql As String
    Dim ErrNum As Long, ErrDesc As String, ErrSrc As String
    Dim ColDpto As String, Corte As String, Cols As String

    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open Conexion
    On Error GoTo Falla

    If ConAnterior Then
        If EnDolares Then ValidaTasas cn, DateSerial(1900, 1, 1), FechaHasta, TiposCuenta
    Else
        If EnDolares Then ValidaTasas cn, FechaDesde, FechaHasta, TiposCuenta
    End If

    If PorDpto Then
        ColDpto = SqlDpto(False)
    Else
        ColDpto = "'-'"
    End If

    'ConAnterior: columna 1 = todo lo anterior a FechaDesde, columna 2 = FechaDesde..FechaHasta.
    If ConAnterior Then
        Corte = "'" & SqlFecha(FechaDesde) & "'"
        Cols = "SUM(CASE WHEN t.FechaTransaccion < " & Corte & " THEN " & SqlImporte("Debito", EnDolares) & " END) AS Debito, " & _
               "SUM(CASE WHEN t.FechaTransaccion < " & Corte & " THEN " & SqlImporte("Credito", EnDolares) & " END) AS Credito, " & _
               "SUM(CASE WHEN t.FechaTransaccion >= " & Corte & " THEN " & SqlImporte("Debito", EnDolares) & " END) AS Debito2, " & _
               "SUM(CASE WHEN t.FechaTransaccion >= " & Corte & " THEN " & SqlImporte("Credito", EnDolares) & " END) AS Credito2 "
    Else
        Cols = "SUM(" & SqlImporte("Debito", EnDolares) & ") AS Debito, " & _
               "SUM(" & SqlImporte("Credito", EnDolares) & ") AS Credito, " & _
               "CAST(0 AS float) AS Debito2, CAST(0 AS float) AS Credito2 "
    End If

    Sql = "SELECT c.CodCuentas, c.DescripcionCuentas, c.KeyGrupo, c.TipoCuenta, " & ColDpto & " AS CodDepartamento, " & Cols & ", c.UbicacionReporte " & _
          SqlDesdeDonde(FechaDesde, FechaHasta, EnDolares, ExcluirCierre, DptoDesde, DptoHasta, TiposCuenta, ConAnterior, AjusteExcluir) & _
          "GROUP BY c.CodCuentas, c.DescripcionCuentas, c.KeyGrupo, c.TipoCuenta, c.UbicacionReporte" & IIf(PorDpto, ", " & ColDpto, "") & " " & _
          "ORDER BY c.KeyGrupo, " & IIf(PorDpto, ColDpto & ", ", "") & "c.CodCuentas"

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
'   ExcluirAjuste True = excluye los asientos de ajuste de la otra moneda (como los reportes acumulados)
'   Balance       True = Balance General: cuentas A/B/C, "Resultado Periodo" dentro de Capital, Total Pasivo
'                 mas Capital y Diferencial Cambiario (lo usa RptBalanceGenerar)
' Devuelve un Recordset desconectado listo para ArepResultado / ArepBalance...
'-----------------------------------------------------------------------------------------------------
Public Function RptResultadoDptoGenerar(ByVal FechaDesde As Date, ByVal FechaHasta As Date, _
        ByVal EnDolares As Boolean, ByVal ExcluirCierre As Boolean, _
        ByVal DptoDesde As String, ByVal DptoHasta As String, _
        ByVal NivelMax As Integer, ByVal OcultarCeros As Boolean, _
        Optional ByRef Advertencias As String = "", _
        Optional ByVal Prefijos As String = "D,G,O", _
        Optional ByVal PorDpto As Boolean = True, Optional ByVal ConAnterior As Boolean = False, _
        Optional ByVal ExcluirAjuste As Boolean = False, Optional ByVal Balance As Boolean = False) As ADODB.Recordset
    Dim rsSal As ADODB.Recordset, rsRes As ADODB.Recordset, dKeys As Object
    Dim i As Long, g As Long, Clave As String, Tipos As String, Ajuste As String
    Dim IngresosTot As Double, GastosTot As Double, SinGrupo As Double
    Dim IngresosTot2 As Double, GastosTot2 As Double
    Dim TotActivo1 As Double, TotPC1 As Double, TotPC2 As Double, Dif As Double

    Advertencias = ""
    mNivelMax = NivelMax
    mOcultarCeros = OcultarCeros
    mPorDpto = PorDpto
    mConAnterior = ConAnterior
    mBalance = Balance
    mResul1 = 0
    mResul2 = 0
    mOrden = 0

    If Balance Then
        Tipos = TIPOS_BALANCE
    Else
        Tipos = TIPOS_RESULTADO
    End If
    If ExcluirAjuste Then
        If EnDolares Then Ajuste = "C" & ChrW(243) & "rdobas" Else Ajuste = "D" & ChrW(243) & "lares"
    End If

    Set rsSal = RptSaldosPorCuentaDpto(FechaDesde, FechaHasta, EnDolares, ExcluirCierre, DptoDesde, DptoHasta, _
                                       Tipos, PorDpto, ConAnterior, Ajuste)
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
        ReDim mMonto2(0 To mNSal - 1)
        ReDim mSalGrp(0 To mNSal - 1)
    Else
        ReDim mMonto(0 To 0)
        ReDim mMonto2(0 To 0)
        ReDim mSalGrp(0 To 0)
    End If

    For i = 0 To mNSal - 1
        If EsNaturalezaCredito(mSal(3, i) & "") Then
            mMonto(i) = Dbl(mSal(6, i)) - Dbl(mSal(5, i))
            mMonto2(i) = Dbl(mSal(8, i)) - Dbl(mSal(7, i))
            IngresosTot = IngresosTot + mMonto(i)
            IngresosTot2 = IngresosTot2 + mMonto2(i)
        Else
            mMonto(i) = Dbl(mSal(5, i)) - Dbl(mSal(6, i))
            mMonto2(i) = Dbl(mSal(7, i)) - Dbl(mSal(8, i))
            GastosTot = GastosTot + mMonto(i)
            GastosTot2 = GastosTot2 + mMonto2(i)
        End If
        Clave = mSal(2, i) & ""
        If dKeys.Exists(Clave) Then
            g = dKeys(Clave)
            mSalGrp(i) = g
            mGrp(g).Total = mGrp(g).Total + mMonto(i)
            mGrp(g).Total2 = mGrp(g).Total2 + mMonto2(i)
            If Abs(mMonto(i)) > EPS Or Abs(mMonto2(i)) > EPS Then mGrp(g).Datos = True
        Else
            mSalGrp(i) = -1
            SinGrupo = SinGrupo + mMonto(i)
            Advertencias = Advertencias & "Cuenta " & mSal(0, i) & " (grupo '" & Clave & "', dpto " & mSal(4, i) & _
                           ") tiene saldo " & Format$(mMonto(i), "#,##0.00") & " y no esta en la estructura." & vbLf
        End If
    Next

    '--- Balance: el resultado (ingresos - costos y gastos) se suma al grupo C (Capital) y se muestra
    '    como la fila "Resultado Periodo". Columna 1 = acumulado anterior, columna 2 = periodo.
    If Balance Then
        Set rsRes = RptSaldosPorCuentaDpto(FechaDesde, FechaHasta, EnDolares, ExcluirCierre, "", "", _
                                           TIPOS_RESULTADO, False, True, Ajuste)
        Do While Not rsRes.EOF
            If (rsRes.Fields("TipoCuenta").Value & "") = TIPO_INGRESOS Then
                mResul1 = mResul1 + Dbl(rsRes.Fields("Credito").Value) - Dbl(rsRes.Fields("Debito").Value)
                mResul2 = mResul2 + Dbl(rsRes.Fields("Credito2").Value) - Dbl(rsRes.Fields("Debito2").Value)
            Else
                mResul1 = mResul1 - (Dbl(rsRes.Fields("Debito").Value) - Dbl(rsRes.Fields("Credito").Value))
                mResul2 = mResul2 - (Dbl(rsRes.Fields("Debito2").Value) - Dbl(rsRes.Fields("Credito2").Value))
            End If
            rsRes.MoveNext
        Loop
        rsRes.Close
        If dKeys.Exists("C") Then
            g = dKeys("C")
            mGrp(g).Total = mGrp(g).Total + mResul1
            mGrp(g).Total2 = mGrp(g).Total2 + mResul2
            If Abs(mResul1) > EPS Or Abs(mResul2) > EPS Then mGrp(g).Datos = True
        ElseIf Abs(mResul1) > EPS Or Abs(mResul2) > EPS Then
            Advertencias = Advertencias & "No existe el grupo C (Capital): el Resultado Periodo no se pudo incluir." & vbLf
        End If
    End If

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

    If Balance Then
        'Total Pasivo + Capital = grupos B y C. AjusteDiferencial del original: si Activo - (Pasivo + Capital)
        'en la columna 1 es menor a 0.02 se agrega "Diferencial Cambiario" y se suma al total.
        If dKeys.Exists("A") Then TotActivo1 = mGrp(dKeys("A")).Total
        If dKeys.Exists("B") Then
            TotPC1 = TotPC1 + mGrp(dKeys("B")).Total
            TotPC2 = TotPC2 + mGrp(dKeys("B")).Total2
        End If
        If dKeys.Exists("C") Then
            TotPC1 = TotPC1 + mGrp(dKeys("C")).Total
            TotPC2 = TotPC2 + mGrp(dKeys("C")).Total2
        End If
        Dif = Round(Round(TotActivo1, 2) - Round(TotPC1, 2), 2)
        If Abs(Dif) < 0.02 And Abs(Dif) <> 0 Then
            AgregaFila "Diferencial Cambiario", "DF", "", "", 1, 0, Dif, 0, 0, True, "", ""
            TotPC1 = TotPC1 + Dif
            TotPC2 = TotPC2 + Dif
        End If
        AgregaFila " Total Pasivo M" & ChrW(224) & "s Capital", "PC", "", "", 1, 0, TotPC1, 0, TotPC2, True, "", ""
    Else
        AgregaFila " Total Costos y Gastos", "CG", "", "", 1, 0, GastosTot, 0, GastosTot2, True, "", ""
        AgregaFila " Resultado Periodo", "RP", "", "", 1, 0, IngresosTot - GastosTot, 0, IngresosTot2 - GastosTot2, True, "", ""
    End If

    If SinGrupo <> 0 Then
        Advertencias = Advertencias & "El total incluye " & Format$(SinGrupo, "#,##0.00") & _
                       " de cuentas que no aparecen en la estructura." & vbLf
    End If

    If mRs.RecordCount > 0 Then mRs.MoveFirst
    Set RptResultadoDptoGenerar = mRs
    Set mRs = Nothing
    mSal = Empty
End Function

Private Function EsNaturalezaCredito(ByVal Tipo As String) As Boolean
    Select Case Tipo
        Case TIPO_INGRESOS, "Otros Pasivos", "Cuentas x Pagar", "Pasivo", "Capital"
            EsNaturalezaCredito = True
    End Select
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
                mGrp(g).Total2 = mGrp(g).Total2 + mGrp(c).Total2
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
               mGrp(g).Nivel, 0, 0, 0, 0, False, "", ""

    EmitirCuentasGrupo g

    For c = 0 To mNGrp - 1
        If c <> g Then
            If mGrp(c).Superior = mGrp(g).Key Then EmitirGrupo c
        End If
    Next

    'Balance: la utilidad del ejercicio va como una cuenta mas del Capital, justo antes de su total
    If mBalance And StrComp(mGrp(g).Key, "C", vbTextCompare) = 0 Then
        mFilaUbic = "Resultado Periodo"
        AgregaFila Space$(mGrp(g).Nivel) & " Resultado Periodo", "RP", "", mGrp(g).Key, mGrp(g).Nivel, _
                   mResul1, 0, mResul2, 0, False, "", ""
    End If

    AgregaFila Space$(mGrp(g).Nivel - 1) & "Total " & mGrp(g).Descripcion, mGrp(g).Key, mGrp(g).Superior, "", _
               mGrp(g).Nivel, 0, mGrp(g).Total, 0, mGrp(g).Total2, True, "", ""
End Sub

' Cuentas del grupo. Con departamentos: encabezado, cuentas y total por departamento.
' Sin departamentos (mPorDpto = False): solo las cuentas, en orden de codigo.
Private Sub EmitirCuentasGrupo(ByVal g As Long)
    Dim i As Long, Dpto As String, DptoActual As String, Abierto As Boolean, TotDpto As Double, TotDpto2 As Double
    Dim Nivel As Integer, Key As String, Sup As String

    Nivel = mGrp(g).Nivel
    Key = mGrp(g).Key
    Sup = mGrp(g).Superior

    For i = 0 To mNSal - 1
        If mSalGrp(i) = g Then
            If Not (mOcultarCeros And Abs(mMonto(i)) < EPS And Abs(mMonto2(i)) < EPS) Then
                Dpto = mSal(4, i) & ""
                If mPorDpto Then
                    If Abierto And Dpto <> DptoActual Then
                        AgregaFila Space$(Nivel) & "Total " & NombreDpto(DptoActual), Key, Sup, "", Nivel + 1, 0, TotDpto, 0, TotDpto2, True, DptoActual, ""
                        Abierto = False
                    End If
                    If Not Abierto Then
                        AgregaFila Space$(Nivel) & NombreDpto(Dpto), Key, Sup, "", Nivel + 1, 0, 0, 0, 0, False, Dpto, ""
                        DptoActual = Dpto
                        TotDpto = 0
                        TotDpto2 = 0
                        Abierto = True
                    End If
                End If
                mFilaUbic = UbicacionDeCuenta(mSal(9, i) & "", mSal(3, i) & "")
                mFilaTipo = mSal(3, i) & ""
                mFilaDescCta = mSal(1, i) & ""
                AgregaFila Space$(Nivel) & mSal(0, i) & "." & (mSal(1, i) & ""), mSal(0, i) & "", "", Key, Nivel + 1, _
                           mMonto(i), 0, mMonto2(i), 0, False, Dpto, mSal(0, i) & ""
                TotDpto = TotDpto + mMonto(i)
                TotDpto2 = TotDpto2 + mMonto2(i)
            End If
        End If
    Next
    If Abierto Then
        AgregaFila Space$(Nivel) & "Total " & NombreDpto(DptoActual), Key, Sup, "", Nivel + 1, 0, TotDpto, 0, TotDpto2, True, DptoActual, ""
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

' Agrega una fila aplicando el filtro de nivel (misma regla que CmbNivel en el reporte original:
Private Sub AgregaFila(ByVal Descripcion As String, ByVal KeyGrupo As String, ByVal KeySuperior As String, _
        ByVal KeyCuenta As String, ByVal Nivel As Integer, ByVal Debe1 As Double, ByVal Haber1 As Double, _
        ByVal Debe2 As Double, ByVal Haber2 As Double, _
        ByVal EsTotal As Boolean, ByVal CodDpto As String, ByVal CodCuenta As String)
    If mNivelMax > 0 Then
        If Nivel > mNivelMax Or (Nivel = mNivelMax And Not EsTotal) Then
            mFilaUbic = "": mFilaTipo = "": mFilaDescCta = ""
            Exit Sub
        End If
    End If
    mOrden = mOrden + 1
    mRs.AddNew
    mRs.Fields("Descripcion").Value = Descripcion
    mRs.Fields("Debe1").Value = Debe1
    mRs.Fields("Haber1").Value = Haber1
    mRs.Fields("Debe2").Value = Debe2
    mRs.Fields("Haber2").Value = Haber2
    'Columna 3 = columna 1 + columna 2 (igual que "Update Reportes set debe3=debe1+debe2,haber3=haber1+haber2");
    'el reporte por departamento nunca la lleno, se deja en 0.
    If mPorDpto Then
        mRs.Fields("Debe3").Value = 0
        mRs.Fields("Haber3").Value = 0
    Else
        mRs.Fields("Debe3").Value = Debe1 + Debe2
        mRs.Fields("Haber3").Value = Haber1 + Haber2
    End If
    mRs.Fields("TotalDebe").Value = Debe1 + Debe2
    mRs.Fields("TotalHaber").Value = Haber1 + Haber2
    mRs.Fields("KeyGrupo").Value = KeyGrupo
    mRs.Fields("KeyGrupoSuperior").Value = KeySuperior
    mRs.Fields("KeyGrupoCuenta").Value = KeyCuenta
    mRs.Fields("Nivel").Value = Nivel
    mRs.Fields("Orden").Value = mOrden
    mRs.Fields("CodDepartamento").Value = CodDpto
    mRs.Fields("CodCuentas").Value = CodCuenta
    mRs.Fields("Ubicacion").Value = mFilaUbic
    mRs.Fields("TipoCuenta").Value = mFilaTipo
    mRs.Fields("DescripcionCuentas").Value = mFilaDescCta
    mFilaUbic = ""
    mFilaTipo = ""
    mFilaDescCta = ""
    mRs.Update
End Sub


'-----------------------------------------------------------------------------------------------------
' Atajos y utilidades para los reportes sin departamento
'-----------------------------------------------------------------------------------------------------
' ESTADO DE RESULTADO (ConAnterior = False) y RESULTADO ACUMULADO / HISTORICO (ConAnterior = True:
' columna 1 = actividad del periodo, columna 2 = saldo anterior a FechaDesde, columna 3 = suma).
Public Function RptResultadoGenerar(ByVal FechaDesde As Date, ByVal FechaHasta As Date, _
        ByVal EnDolares As Boolean, ByVal ExcluirCierre As Boolean, _
        ByVal NivelMax As Integer, ByVal OcultarCeros As Boolean, ByVal ConAnterior As Boolean, _
        Optional ByRef Advertencias As String = "") As ADODB.Recordset
    'Los reportes acumulados originales (SaldoReportesAcumulado) siempre excluyen los ajustes de la otra moneda
    Set RptResultadoGenerar = RptResultadoDptoGenerar(FechaDesde, FechaHasta, EnDolares, ExcluirCierre, "", "", _
                                  NivelMax, OcultarCeros, Advertencias, "D,G,O", False, ConAnterior, ConAnterior)
End Function

' BALANCE GENERAL / ACUMULADO / HISTORICO. Columna 1 = saldo anterior a FechaDesde, columna 2 = actividad
' del periodo, columna 3 = saldo a FechaHasta (suma). Incluye Resultado Periodo en Capital, Total Pasivo
' mas Capital y el Diferencial Cambiario (AjusteDiferencial) cuando corresponde.
Public Function RptBalanceGenerar(ByVal FechaDesde As Date, ByVal FechaHasta As Date, _
        ByVal EnDolares As Boolean, ByVal NivelMax As Integer, ByVal OcultarCeros As Boolean, _
        Optional ByRef Advertencias As String = "") As ADODB.Recordset
    Set RptBalanceGenerar = RptResultadoDptoGenerar(FechaDesde, FechaHasta, EnDolares, False, "", "", _
                                NivelMax, OcultarCeros, Advertencias, "A,B,C", False, True, True, True)
End Function

' Divide todas las columnas del reporte entre la tasa (equivale a ConvertirReporte, sin tocar la BD).
Public Sub RptConvertirRecordset(rs As ADODB.Recordset, ByVal TasaCambio As Double)
    Dim Campo As Variant
    If TasaCambio = 0 Then Err.Raise vbObjectError + 1002, "ResultadoDpto", "La tasa de cambio no puede ser cero."
    If rs.BOF And rs.EOF Then Exit Sub
    rs.MoveFirst
    Do While Not rs.EOF
        For Each Campo In VBA.Split("Debe1,Haber1,Debe2,Haber2,Debe3,Haber3,TotalDebe,TotalHaber", ",")
            rs.Fields(CStr(Campo)).Value = Round(CDbl(rs.Fields(CStr(Campo)).Value) / TasaCambio, 2)
        Next
        rs.Update
        rs.MoveNext
    Loop
    rs.MoveFirst
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

' True si la estructura trae los grupos de Utilidad Bruta configurados (MDIPrimero.AdoConfiguracion);
' en ese caso el reporte original reordena filas con Utilidadbruta, que esta version todavia no hace.
Public Function RptHayUtilidadBruta(rs As ADODB.Recordset) As Boolean
    Dim Nombres As Variant, N As Variant, Texto As String, Hay As Boolean
    With MDIPrimero.AdoConfiguracion.Recordset
        Nombres = Array(.Fields("OtrosINgresos").Value & "", .Fields("Costos").Value & "", _
                        .Fields("Ingresos").Value & "", .Fields("CostosOperativos").Value & "")
    End With
    For Each N In Nombres
        If Len(Trim$(N)) = 0 Then Exit Function
        Hay = False
        If Not (rs.BOF And rs.EOF) Then
            rs.MoveFirst
            Do While Not rs.EOF
                Texto = rs.Fields("Descripcion").Value & ""
                If InStr(1, Texto, CStr(N), vbTextCompare) > 0 Then Hay = True: Exit Do
                rs.MoveNext
            Loop
            rs.MoveFirst
        End If
        If Not Hay Then Exit Function
    Next
    RptHayUtilidadBruta = True
End Function


'-----------------------------------------------------------------------------------------------------
' BALANCE GENERAL RESUMEN / RESUMEN ANEXOS (en memoria)
' ArepBalancePersonalizado pide cada renglon con SaldosPersonalizados("Cajas"...) y ArepAnexosBalances
' lee las cuentas de Reportes. Mientras RptUbicActivo = True ambos leen de aqui (ver Funciones.bas y
' ArepAnexosBalances). Siempre llamar RptUbicacionesLiberar al terminar.
'-----------------------------------------------------------------------------------------------------
' Ubicacion de la cuenta en el reporte resumen: Cuentas.UbicacionReporte o, si esta vacia, la que
' corresponde al tipo de cuenta (mismo mapeo que CreaEstructura).
Private Function UbicacionDeCuenta(ByVal Ubicacion As String, ByVal TipoCuenta As String) As String
    If Len(Trim$(Ubicacion)) > 0 Then
        UbicacionDeCuenta = Ubicacion
        Exit Function
    End If
    Select Case TipoCuenta
        Case "Caja": UbicacionDeCuenta = "Cajas"
        Case "Bancos": UbicacionDeCuenta = "Bancos"
        Case "Cuentas x Cobrar": UbicacionDeCuenta = "Cuentas x Cobrar"
        Case "Inventario": UbicacionDeCuenta = "Inventario"
        Case "Activo Fijo": UbicacionDeCuenta = "Terreno y Edificios"
        Case "Papeleria - Utiles": UbicacionDeCuenta = "Papeleria y Utiles de Oficina"
        Case "Otros Activos": UbicacionDeCuenta = "Otros Activos"
        Case "Cuentas x Pagar": UbicacionDeCuenta = "Proveedores"
        Case "Pasivo": UbicacionDeCuenta = "Pasivos Acumulados"
        Case "Otros Pasivos": UbicacionDeCuenta = "Otros Pasivos"
        Case "Capital": UbicacionDeCuenta = "Acciones Comunes"
        Case "Ingresos - Ventas": UbicacionDeCuenta = "Ingresos - Ventas"
        Case "Costos": UbicacionDeCuenta = "Costos"
        Case "Gastos": UbicacionDeCuenta = "Gastos"
    End Select
End Function

' Genera el balance completo (sin filtros de nivel ni ceros, como el original) y lo deja disponible para
' los reportes resumen. Devuelve el mismo recordset que RptBalanceGenerar.
Public Function RptBalanceResumenGenerar(ByVal FechaDesde As Date, ByVal FechaHasta As Date, _
        ByVal EnDolares As Boolean, ByVal Convertir As Boolean, _
        Optional ByRef Advertencias As String = "") As ADODB.Recordset
    Dim rs As ADODB.Recordset
    RptUbicacionesLiberar
    Set rs = RptBalanceGenerar(FechaDesde, FechaHasta, EnDolares, 0, False, Advertencias)
    'Resumen Anexos: el original ejecuta ConvertirReporte(FechaFin) siempre (divide todo entre la tasa de FechaFin)
    If Convertir Then RptConvertirRecordset rs, BuscaTasaCambio(FechaHasta)
    RptCargarUbicaciones rs
    Set mRsAnexos = rs
    RptUbicActivo = True
    Set RptBalanceResumenGenerar = rs
End Function

' Equivale a SaldosPersonalizados: SUM(Debe3)-SUM(Haber3) y SUM(Debe2)-SUM(Haber2) de las filas por Ubicacion.
Private Sub RptCargarUbicaciones(rs As ADODB.Recordset)
    Dim Ub As String
    Set mUbicSaldo = CreateObject("Scripting.Dictionary")
    mUbicSaldo.CompareMode = vbTextCompare
    Set mUbicPeriodo = CreateObject("Scripting.Dictionary")
    mUbicPeriodo.CompareMode = vbTextCompare
    mRP2 = 0
    mRP3 = 0
    If rs.BOF And rs.EOF Then Exit Sub
    rs.MoveFirst
    Do While Not rs.EOF
        If (rs.Fields("KeyGrupo").Value & "") = "RP" Then
            mRP3 = CDbl(rs.Fields("Haber3").Value) - CDbl(rs.Fields("Debe3").Value)
            mRP2 = CDbl(rs.Fields("Haber2").Value) - CDbl(rs.Fields("Debe2").Value)
        End If
        Ub = rs.Fields("Ubicacion").Value & ""
        If Len(Ub) > 0 Then
            mUbicSaldo(Ub) = mUbicSaldo(Ub) + CDbl(rs.Fields("Debe3").Value) - CDbl(rs.Fields("Haber3").Value)
            mUbicPeriodo(Ub) = mUbicPeriodo(Ub) + CDbl(rs.Fields("Debe2").Value) - CDbl(rs.Fields("Haber2").Value)
        End If
        rs.MoveNext
    Loop
    rs.MoveFirst
End Sub

' ESTADO DE RESULTADO RESUMEN / RESUMEN 2 / RESUMEN ANEXOS / TRADICIONAL. ArepResultadoResumen* piden cada
' renglon con SaldosPersonalizados y la utilidad con la fila RP de Reportes (RptResultadoRP); mientras
' RptUbicActivo = True leen de aqui. OcultarCeros solo importa para los anexos (lista de cuentas).
Public Function RptResultadoResumenGenerar(ByVal FechaDesde As Date, ByVal FechaHasta As Date, _
        ByVal EnDolares As Boolean, ByVal ExcluirCierre As Boolean, ByVal OcultarCeros As Boolean, _
        Optional ByRef Advertencias As String = "") As ADODB.Recordset
    Dim rs As ADODB.Recordset
    RptUbicacionesLiberar
    Set rs = RptResultadoGenerar(FechaDesde, FechaHasta, EnDolares, ExcluirCierre, 0, OcultarCeros, True, Advertencias)
    RptCargarUbicaciones rs
    Set mRsAnexos = rs
    RptUbicActivo = True
    Set RptResultadoResumenGenerar = rs
End Function

' Utilidad (fila RP): Periodo = True actividad del periodo (Haber2 - Debe2), False saldo (Haber3 - Debe3).
Public Function RptResultadoRP(ByVal Periodo As Boolean) As Double
    If Periodo Then RptResultadoRP = mRP2 Else RptResultadoRP = mRP3
End Function

Public Function RptUbicacionSaldo(ByVal Ubicacion As String) As Double
    If mUbicSaldo.Exists(Ubicacion) Then RptUbicacionSaldo = mUbicSaldo(Ubicacion)
End Function

Public Function RptUbicacionPeriodo(ByVal Ubicacion As String) As Double
    If mUbicPeriodo.Exists(Ubicacion) Then RptUbicacionPeriodo = mUbicPeriodo(Ubicacion)
End Function

' Filas de cuenta para ArepAnexosBalances. Condiciones es el WHERE que arma el reporte con los checks
' (... Reportes.Ubicacion = N'Cajas' OR ...); de ahi se toman los nombres de Ubicacion.
Public Function RptAnexosFiltrar(ByVal Condiciones As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset, Filtro As String, p As Long, q As Long, Nom As String
    Set rs = mRsAnexos.Clone
    p = InStr(Condiciones, "N'")
    Do While p > 0
        q = InStr(p + 2, Condiciones, "'")
        If q = 0 Then Exit Do
        Nom = Mid$(Condiciones, p + 2, q - p - 2)
        If Len(Filtro) > 0 Then Filtro = Filtro & " OR "
        Filtro = Filtro & "Ubicacion = '" & SqlTexto(Nom) & "'"
        p = InStr(q + 1, Condiciones, "N'")
    Loop
    'Solo las filas de cuenta traen Ubicacion; ADO no admite comparar contra '' en Filter
    If Len(Filtro) > 0 Then rs.Filter = Filtro
    rs.Sort = "Orden"
    Set RptAnexosFiltrar = rs
End Function

Public Sub RptUbicacionesLiberar()
    RptUbicActivo = False
    Set mUbicSaldo = Nothing
    Set mUbicPeriodo = Nothing
    Set mRsAnexos = Nothing
End Sub
