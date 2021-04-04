Attribute VB_Name = "CodigoRetirado"
'//////////////////////////////////////////////////////////////
'ESTE CODIGO SE RETIRO DE LA FUNCION SALDO REPORTES, EN LOS SALDOS
'DEL PERIODO ANTERIOR, PARA REMPLAZARLO POR UNA CONSULTA QUE
'HAGA EL CALCULO SUMANDO TODO.



'    Do While Not FrmReportes.DtaConsulta.Recordset.EOF
'
'       TotalDebito = 0
'       TotalCredito = 0
'      TipoCuenta = FrmReportes.DtaConsulta.Recordset("TipoCuenta")
'      TipoMoneda = FrmReportes.DtaConsulta.Recordset("TipoMoneda")
'      FechaTransaccion = FrmReportes.DtaConsulta.Recordset("FechaTransaccion")
'      Fechas1 = FrmReportes.DtaConsulta.Recordset("FechaTransaccion")
'      TasaCambio = FrmReportes.DtaConsulta.Recordset("MontoCordobas")
'     If TasaCambio = 0 Then
'      Cadena = "La tasa de Cambio de Cambio con Fecha: " & Fechas1 & vbLf
'      Cadena = Cadena & "no puede ser igual a Cero, el Sistema Contable" & vbLf
'      Cadena = Cadena & "no contiuara el proceso......"
'      MsgBox Cadena, vbCritical, "Sistema Contable"
'      Exit Sub
'     End If
'      Descripcion = FrmReportes.DtaConsulta.Recordset("CodCuentas") + "." + FrmReportes.DtaConsulta.Recordset("DescripcionCuentas")
'      If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
'        If Not IsNull(FrmReportes.DtaConsulta.Recordset("MDebito")) Then
'            Debito = FrmReportes.DtaConsulta.Recordset("MDebito")
'        End If
'        If Not IsNull(FrmReportes.DtaConsulta.Recordset("MCredito")) Then
'            Credito = FrmReportes.DtaConsulta.Recordset("MCredito")
'        End If
'        Total1 = Debito - Credito + Total1
'
'        '//////////////////Verifico el tipo de moneda de la cuenta////////////////////
'        If Not TipoMoneda = FrmReportes.CmbMoneda.Text Then
'           Select Case TipoMoneda
'              Case "Córdobas"
'                    TotalCuenta = (Debito - Credito) / TasaCambio + TotalCuenta
'
'              Case "Dólares"
'                    TotalCuenta = (Debito - Credito) * TasaCambio + TotalCuenta
'
'          End Select
'        Else
'               TotalCuenta = (Debito - Credito) + TotalCuenta
'        End If
'
'        Debito = 0
'        Credito = 0
'      Else
'         If Not IsNull(FrmReportes.DtaConsulta.Recordset("MDebito")) Then
'            Debito = FrmReportes.DtaConsulta.Recordset("MDebito")
'         End If
'         If Not IsNull(FrmReportes.DtaConsulta.Recordset("MCredito")) Then
'            Credito = FrmReportes.DtaConsulta.Recordset("MCredito")
'         End If
'
'        '//////////////////Verifico el tipo de moneda de la cuenta////////////////////
'        If Not TipoMoneda = FrmReportes.CmbMoneda.Text Then
'           Select Case TipoMoneda
'              Case "Córdobas"
'                    TotalCuenta = (Credito - Debito) / TasaCambio + TotalCuenta
'
'              Case "Dólares"
'                    TotalCuenta = (Credito - Debito) * TasaCambio + TotalCuenta
'
'          End Select
'        Else
'               TotalCuenta = (Credito - Debito) + TotalCuenta
'        End If
'
'         Total1 = Credito - Debito + Total1
'         Debito = 0
'         Credito = 0
'      End If
'
'
'
'
'
'
'
'   FrmReportes.DtaConsulta.Recordset.MoveNext
'
'   Loop
