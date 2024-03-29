/*
   sábado, 16 de marzo de 201309:27:44 a.m.
   Usuario: 
   Servidor: JUAN\SQL2005
   Base de datos: SistemaContableSystems
   Aplicación: 
*/

/* Para evitar posibles problemas de pérdida de datos, debe revisar esta secuencia de comandos detalladamente antes de ejecutarla fuera del contexto del diseñador de base de datos.*/
BEGIN TRANSACTION
SET QUOTED_IDENTIFIER ON
SET ARITHABORT ON
SET NUMERIC_ROUNDABORT OFF
SET CONCAT_NULL_YIELDS_NULL ON
SET ANSI_NULLS ON
SET ANSI_PADDING ON
SET ANSI_WARNINGS ON
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.Transacciones
	DROP CONSTRAINT DF_Transacciones_DescripcionMovimiento
GO
ALTER TABLE dbo.Transacciones
	DROP CONSTRAINT DF_Transacciones_Debito
GO
ALTER TABLE dbo.Transacciones
	DROP CONSTRAINT DF_Transacciones_Credito
GO
ALTER TABLE dbo.Transacciones
	DROP CONSTRAINT DF_Transacciones_Conciliada
GO
ALTER TABLE dbo.Transacciones
	DROP CONSTRAINT DF_Transacciones_ConciliacionProcesada
GO
ALTER TABLE dbo.Transacciones
	DROP CONSTRAINT DF_Transacciones_FechaDescuento
GO
ALTER TABLE dbo.Transacciones
	DROP CONSTRAINT DF_Transacciones_DescuentoDisponible
GO
ALTER TABLE dbo.Transacciones
	DROP CONSTRAINT DF_Transacciones_FechaVence
GO
ALTER TABLE dbo.Transacciones
	DROP CONSTRAINT DF_Transacciones_Beneficiario
GO
ALTER TABLE dbo.Transacciones
	DROP CONSTRAINT DF_Transacciones_CodCuentaProveedor
GO
ALTER TABLE dbo.Transacciones
	DROP CONSTRAINT DF_Transacciones_TipoFactura
GO
ALTER TABLE dbo.Transacciones
	DROP CONSTRAINT DF_Transacciones_DebitoD
GO
ALTER TABLE dbo.Transacciones
	DROP CONSTRAINT DF_Transacciones_CreditoD
GO
CREATE TABLE dbo.Tmp_Transacciones
	(
	CodCuentas nvarchar(50) NOT NULL,
	FechaTransaccion smalldatetime NOT NULL,
	NPeriodo int NOT NULL,
	NTransaccion int NOT NULL IDENTITY (1, 1),
	NumeroMovimiento int NULL,
	NombreCuenta nvarchar(255) NULL,
	VoucherNo nvarchar(50) NULL,
	DescripcionMovimiento nvarchar(255) NULL,
	Clave nvarchar(10) NULL,
	TCambio float(53) NULL,
	Debito decimal(18, 2) NULL,
	Credito decimal(18, 2) NULL,
	FacturaNo nvarchar(50) NULL,
	ChequeNo nvarchar(50) NULL,
	Fuente nvarchar(50) NULL,
	FechaTasas smalldatetime NULL,
	Conciliada int NULL,
	ConciliacionProcesada int NULL,
	FechaDescuento smalldatetime NULL,
	DescuentoDisponible money NULL,
	FechaVence smalldatetime NULL,
	Beneficiario nvarchar(255) NULL,
	CodCuentaProveedor nvarchar(50) NULL,
	TipoFactura nvarchar(50) NULL,
	DebitoD decimal(18, 2) NULL,
	CreditoD decimal(18, 2) NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.Tmp_Transacciones ADD CONSTRAINT
	DF_Transacciones_DescripcionMovimiento DEFAULT (N'*') FOR DescripcionMovimiento
GO
ALTER TABLE dbo.Tmp_Transacciones ADD CONSTRAINT
	DF_Transacciones_Debito DEFAULT (0) FOR Debito
GO
ALTER TABLE dbo.Tmp_Transacciones ADD CONSTRAINT
	DF_Transacciones_Credito DEFAULT (0) FOR Credito
GO
ALTER TABLE dbo.Tmp_Transacciones ADD CONSTRAINT
	DF_Transacciones_Conciliada DEFAULT (0) FOR Conciliada
GO
ALTER TABLE dbo.Tmp_Transacciones ADD CONSTRAINT
	DF_Transacciones_ConciliacionProcesada DEFAULT (0) FOR ConciliacionProcesada
GO
ALTER TABLE dbo.Tmp_Transacciones ADD CONSTRAINT
	DF_Transacciones_FechaDescuento DEFAULT (1 / 1 / 1900) FOR FechaDescuento
GO
ALTER TABLE dbo.Tmp_Transacciones ADD CONSTRAINT
	DF_Transacciones_DescuentoDisponible DEFAULT (0) FOR DescuentoDisponible
GO
ALTER TABLE dbo.Tmp_Transacciones ADD CONSTRAINT
	DF_Transacciones_FechaVence DEFAULT (1 / 1 / 1900) FOR FechaVence
GO
ALTER TABLE dbo.Tmp_Transacciones ADD CONSTRAINT
	DF_Transacciones_Beneficiario DEFAULT (N'-') FOR Beneficiario
GO
ALTER TABLE dbo.Tmp_Transacciones ADD CONSTRAINT
	DF_Transacciones_CodCuentaProveedor DEFAULT (0) FOR CodCuentaProveedor
GO
ALTER TABLE dbo.Tmp_Transacciones ADD CONSTRAINT
	DF_Transacciones_TipoFactura DEFAULT (N'-') FOR TipoFactura
GO
ALTER TABLE dbo.Tmp_Transacciones ADD CONSTRAINT
	DF_Transacciones_DebitoD DEFAULT ((0)) FOR DebitoD
GO
ALTER TABLE dbo.Tmp_Transacciones ADD CONSTRAINT
	DF_Transacciones_CreditoD DEFAULT ((0)) FOR CreditoD
GO
SET IDENTITY_INSERT dbo.Tmp_Transacciones ON
GO
IF EXISTS(SELECT * FROM dbo.Transacciones)
	 EXEC('INSERT INTO dbo.Tmp_Transacciones (CodCuentas, FechaTransaccion, NPeriodo, NTransaccion, NumeroMovimiento, NombreCuenta, VoucherNo, DescripcionMovimiento, Clave, TCambio, Debito, Credito, FacturaNo, ChequeNo, Fuente, FechaTasas, Conciliada, ConciliacionProcesada, FechaDescuento, DescuentoDisponible, FechaVence, Beneficiario, CodCuentaProveedor, TipoFactura, DebitoD, CreditoD)
		SELECT CodCuentas, FechaTransaccion, NPeriodo, NTransaccion, NumeroMovimiento, NombreCuenta, VoucherNo, DescripcionMovimiento, Clave, TCambio, CONVERT(decimal(18, 2), Debito), CONVERT(decimal(18, 2), Credito), FacturaNo, ChequeNo, Fuente, FechaTasas, Conciliada, ConciliacionProcesada, FechaDescuento, DescuentoDisponible, FechaVence, Beneficiario, CodCuentaProveedor, TipoFactura, CONVERT(decimal(18, 2), DebitoD), CONVERT(decimal(18, 2), CreditoD) FROM dbo.Transacciones WITH (HOLDLOCK TABLOCKX)')
GO
SET IDENTITY_INSERT dbo.Tmp_Transacciones OFF
GO
DROP TABLE dbo.Transacciones
GO
EXECUTE sp_rename N'dbo.Tmp_Transacciones', N'Transacciones', 'OBJECT' 
GO
ALTER TABLE dbo.Transacciones ADD CONSTRAINT
	PK_Transacciones PRIMARY KEY CLUSTERED 
	(
	CodCuentas,
	FechaTransaccion,
	NPeriodo,
	NTransaccion
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
