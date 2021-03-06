/*
   sábado, 16 de marzo de 201309:28:46 a.m.
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
ALTER TABLE dbo.DetalleMemoriza
	DROP CONSTRAINT FK_DetalleMemoriza_IndiceMemoriza
GO
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.DetalleMemoriza
	DROP CONSTRAINT DF_DetalleMemoriza_DescripcionMovimiento
GO
ALTER TABLE dbo.DetalleMemoriza
	DROP CONSTRAINT DF_DetalleMemoriza_Debito
GO
ALTER TABLE dbo.DetalleMemoriza
	DROP CONSTRAINT DF_DetalleMemoriza_Credito
GO
ALTER TABLE dbo.DetalleMemoriza
	DROP CONSTRAINT DF_DetalleMemoriza_Conciliada
GO
ALTER TABLE dbo.DetalleMemoriza
	DROP CONSTRAINT DF_DetalleMemoriza_ConciliacionProcesada
GO
ALTER TABLE dbo.DetalleMemoriza
	DROP CONSTRAINT DF_DetalleMemoriza_FechaDescuento
GO
ALTER TABLE dbo.DetalleMemoriza
	DROP CONSTRAINT DF_DetalleMemoriza_DescuentoDisponible
GO
ALTER TABLE dbo.DetalleMemoriza
	DROP CONSTRAINT DF_DetalleMemoriza_FechaVence
GO
ALTER TABLE dbo.DetalleMemoriza
	DROP CONSTRAINT DF_DetalleMemoriza_Beneficiario
GO
ALTER TABLE dbo.DetalleMemoriza
	DROP CONSTRAINT DF_DetalleMemoriza_CodCuentaProveedor
GO
ALTER TABLE dbo.DetalleMemoriza
	DROP CONSTRAINT DF_DetalleMemoriza_TipoFactura
GO
CREATE TABLE dbo.Tmp_DetalleMemoriza
	(
	CodCuentas nvarchar(50) NOT NULL,
	FechaTransaccion smalldatetime NOT NULL,
	NPeriodo int NOT NULL,
	NTransaccion int NOT NULL IDENTITY (1, 1),
	IdMemoria int NOT NULL,
	NombreCuenta nvarchar(255) NULL,
	VoucherNo nvarchar(50) NULL,
	DescripcionMovimiento nvarchar(255) NULL,
	Clave nvarchar(10) NULL,
	TCambio float(53) NULL,
	Debito numeric(18, 2) NULL,
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
	TipoFactura nvarchar(50) NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.Tmp_DetalleMemoriza ADD CONSTRAINT
	DF_DetalleMemoriza_DescripcionMovimiento DEFAULT (N'*') FOR DescripcionMovimiento
GO
ALTER TABLE dbo.Tmp_DetalleMemoriza ADD CONSTRAINT
	DF_DetalleMemoriza_Debito DEFAULT (0) FOR Debito
GO
ALTER TABLE dbo.Tmp_DetalleMemoriza ADD CONSTRAINT
	DF_DetalleMemoriza_Credito DEFAULT (0) FOR Credito
GO
ALTER TABLE dbo.Tmp_DetalleMemoriza ADD CONSTRAINT
	DF_DetalleMemoriza_Conciliada DEFAULT (0) FOR Conciliada
GO
ALTER TABLE dbo.Tmp_DetalleMemoriza ADD CONSTRAINT
	DF_DetalleMemoriza_ConciliacionProcesada DEFAULT (0) FOR ConciliacionProcesada
GO
ALTER TABLE dbo.Tmp_DetalleMemoriza ADD CONSTRAINT
	DF_DetalleMemoriza_FechaDescuento DEFAULT (1 / 1 / 1900) FOR FechaDescuento
GO
ALTER TABLE dbo.Tmp_DetalleMemoriza ADD CONSTRAINT
	DF_DetalleMemoriza_DescuentoDisponible DEFAULT (0) FOR DescuentoDisponible
GO
ALTER TABLE dbo.Tmp_DetalleMemoriza ADD CONSTRAINT
	DF_DetalleMemoriza_FechaVence DEFAULT (1 / 1 / 1900) FOR FechaVence
GO
ALTER TABLE dbo.Tmp_DetalleMemoriza ADD CONSTRAINT
	DF_DetalleMemoriza_Beneficiario DEFAULT (N'-') FOR Beneficiario
GO
ALTER TABLE dbo.Tmp_DetalleMemoriza ADD CONSTRAINT
	DF_DetalleMemoriza_CodCuentaProveedor DEFAULT (0) FOR CodCuentaProveedor
GO
ALTER TABLE dbo.Tmp_DetalleMemoriza ADD CONSTRAINT
	DF_DetalleMemoriza_TipoFactura DEFAULT (N'-') FOR TipoFactura
GO
SET IDENTITY_INSERT dbo.Tmp_DetalleMemoriza ON
GO
IF EXISTS(SELECT * FROM dbo.DetalleMemoriza)
	 EXEC('INSERT INTO dbo.Tmp_DetalleMemoriza (CodCuentas, FechaTransaccion, NPeriodo, NTransaccion, IdMemoria, NombreCuenta, VoucherNo, DescripcionMovimiento, Clave, TCambio, Debito, Credito, FacturaNo, ChequeNo, Fuente, FechaTasas, Conciliada, ConciliacionProcesada, FechaDescuento, DescuentoDisponible, FechaVence, Beneficiario, CodCuentaProveedor, TipoFactura)
		SELECT CodCuentas, FechaTransaccion, NPeriodo, NTransaccion, IdMemoria, NombreCuenta, VoucherNo, DescripcionMovimiento, Clave, TCambio, CONVERT(numeric(18, 2), Debito), CONVERT(decimal(18, 2), Credito), FacturaNo, ChequeNo, Fuente, FechaTasas, Conciliada, ConciliacionProcesada, FechaDescuento, DescuentoDisponible, FechaVence, Beneficiario, CodCuentaProveedor, TipoFactura FROM dbo.DetalleMemoriza WITH (HOLDLOCK TABLOCKX)')
GO
SET IDENTITY_INSERT dbo.Tmp_DetalleMemoriza OFF
GO
DROP TABLE dbo.DetalleMemoriza
GO
EXECUTE sp_rename N'dbo.Tmp_DetalleMemoriza', N'DetalleMemoriza', 'OBJECT' 
GO
ALTER TABLE dbo.DetalleMemoriza ADD CONSTRAINT
	PK_DetalleMemoriza PRIMARY KEY CLUSTERED 
	(
	CodCuentas,
	FechaTransaccion,
	NPeriodo,
	NTransaccion
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
ALTER TABLE dbo.DetalleMemoriza ADD CONSTRAINT
	FK_DetalleMemoriza_IndiceMemoriza FOREIGN KEY
	(
	IdMemoria
	) REFERENCES dbo.IndiceMemoriza
	(
	IdMemoria
	) ON UPDATE  CASCADE 
	 ON DELETE  CASCADE 
	
GO
COMMIT
