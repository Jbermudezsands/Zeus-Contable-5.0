Attribute VB_Name = "Variables"
Public Ejecutar As ADODB.Connection, QuienReporte As String, FechaIngreso As Date
Public objExcel As Excel.Application

Public SaldoFinalAuxiliar As Double, TotalDebitoAux As Double, TotalCreditoAux As Double

Public TotalConDebito As Double, TotalConCredito As Double, NumeroFact As String
Public TotalInicial As Double, TotalFinal As Double

Public ResultadoPersonalizado As Double, ResultadoPersonalizadoPeriodo As Double, RutaArchivo As String
Public Cheque As Boolean, Debitos As Double, Creditos As Double
Public ruta As String, RutaLogo As String, RutaIconos As String, RutaFotos As String
Public Conexion As String, ConexionReporte As String, numero As String
Public Unidad As String, CodGrupo As String, MonedaCuenta As String, MonedaContracuenta As String
Public Criterio As String, CodEncargado As String
Public Opt As String, CodigoCuenta As String
Public estado As String, Color As String, Fechas1 As String, Fechas2 As String
Public NTransacciones As Integer, NumFecha2 As Long, NumFecha1 As Long
Public NPeriodo As Integer, NPeriodoFin As Integer
Public FechaIni As Date, FechaFin As Date, TasaCambioCordobas As Double, TasaCambioEuro As Double
Public NumFecha As Long, Tasa As Boolean, TasaCambioDolares As Double
Public NumeroPeriodo As Integer, EstadoPeriodo As String, NumeroTabla As Integer
Public NumeroTransaccion As Integer, Periodo As Integer, NumeroPeriodo1 As Integer, NumeroPeriodo2 As Integer
Public QueProducto As String, DescripcionContracuenta As String
Public Orden As Boolean, Descripcion As String
Public Debito As Double, Credito As Double, Diferencia As Double, DescripcionCuenta As String
Public NombreEmpresa As String, RUC As String


Public ConsultaTotalesMovimientos As String ' consulta que voy a utilizar para poder cambiar la tabla de reportes para la
'balanza de comparacion

Public TotalDebitoH As Double, TotalCreditoH As Double
Public DebitoAnt As Double, CreditoAnt As Double, TotalHistorico As Double
Public TotalDebito As Double, TotalCredito As Double, TotalDiferencia As Double
Public ClaveAnt As String, Clave As String
Public NumeroAnterior As Double, Lectura As String
Public Respuesta As String, Origen As String
Public TipoCuenta As String, SaldoFinal As Double
Public NumeroPeriodoAnterior As Integer, J As Integer
Public DiferenciaSaldo As Double, NivelAcceso As Integer, NombreUsuario As String
Public GravaUsuarios As Boolean, mes As Date
Public Nivel2 As Integer, Permiso As Boolean, Año As Date
Public Fecha As Date, Encontrado As Boolean, Primero As Boolean
Public Cambio As Double, Salir As Boolean, Monto As Double
Public TotalIni As Double, TotalFin As Double
Public CodUsuario As Integer, CodigoUsuario As Integer, MontoAcordado As Double
Public SaldoIni As Double, SaldoFin As Double, Consecutivo As Integer
Public QUIEN As String, TipoMoneda As String
Public NMovimiento As Integer, Indice As Integer
Public Total As Double, Presupuesto As Double
Public MontoTasa As Double, MontoCheque As Double
Public Fecha1 As Date, Fecha2 As Date, CodigoBanco As String
Public NumeroVoucher As Double, ConsecutivoVoucher As Integer
Public KeyGrupoCuenta As String, NtransaccionPeriodo As Integer
Public NodX As Node, MatrizCuentas() As String
Public DescripcionNodo As String, KeyNodoUltimo As String
Public KeyPadre As String, KeyHijo As String
Public Tipo As String, Imagen As Integer, NodoBase As Boolean
Public KeyPrincipal As String, CambiarTipo As Boolean, i As Integer
Public SaldoEstadoCta As Double, SaldoLibros As Double
Public Periodo1 As Integer, Periodo2 As Integer, GranTotal As Double
Public Total1 As Double, Tabla As Integer, CkNo As Integer
Public FechaTransaccion As Long
Public TasaCambio As Double, TotalDebito1 As Double, TotalCredito1 As Double, TotalCuenta As Double
Public KeyAnterior As String, Totalingresos As Double, TotalGastos As Double
Public NombreBD As String, FechaSistema As Date


Public FechaVence As Date, FechaFactura As Date
Public CodigoCuentaIva As String, Beneficiario As String

Public col As TrueOleDBGrid80.Column
Public cols As TrueOleDBGrid80.Columns

Public rsa As New ADODB.Recordset
Public rsa1 As New ADODB.Recordset
Public rsa2 As New ADODB.Recordset
Public rsa3 As New ADODB.Recordset
Public rsa4 As New ADODB.Recordset

Public wfecha As Date 'JUAN
Public wtextc As String 'JUAN
Public wnombre As String  'nombre de la sede
Public wnombreFinca As String  'nombre de la finca
Public wforma As Form 'JUAN

Public Moneda As String, FechaCheque As String, NumeroMovimientos As String


Public RTotalActivoCirculante As Double, RTotalActivoFijo As Double, RTotalPasivo As Double, RTotalCapital As Double, RUtilidad As Double

Public SaldoIniBank As Double, CodigoBancoBank As String, FechaIniBank As Date, FechaFinBank As Date
