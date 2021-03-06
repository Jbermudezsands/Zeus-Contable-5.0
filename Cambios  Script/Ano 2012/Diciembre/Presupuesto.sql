/*
   sábado, 16 de marzo de 201309:29:34 a.m.
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
ALTER TABLE dbo.Presupuesto
	DROP CONSTRAINT FK_Presupuesto_Periodos
GO
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.Presupuesto
	DROP CONSTRAINT FK_Presupuesto_Cuentas
GO
COMMIT
BEGIN TRANSACTION
GO
CREATE TABLE dbo.Tmp_Presupuesto
	(
	NPeriodo int NOT NULL,
	CodCuenta nvarchar(50) NOT NULL,
	MontoPresupuestado decimal(18, 2) NULL,
	SaldoReal decimal(18, 2) NULL
	)  ON [PRIMARY]
GO
IF EXISTS(SELECT * FROM dbo.Presupuesto)
	 EXEC('INSERT INTO dbo.Tmp_Presupuesto (NPeriodo, CodCuenta, MontoPresupuestado, SaldoReal)
		SELECT NPeriodo, CodCuenta, CONVERT(decimal(18, 2), MontoPresupuestado), CONVERT(decimal(18, 2), SaldoReal) FROM dbo.Presupuesto WITH (HOLDLOCK TABLOCKX)')
GO
DROP TABLE dbo.Presupuesto
GO
EXECUTE sp_rename N'dbo.Tmp_Presupuesto', N'Presupuesto', 'OBJECT' 
GO
ALTER TABLE dbo.Presupuesto ADD CONSTRAINT
	PK_Presupuesto PRIMARY KEY CLUSTERED 
	(
	NPeriodo,
	CodCuenta
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
ALTER TABLE dbo.Presupuesto WITH NOCHECK ADD CONSTRAINT
	FK_Presupuesto_Cuentas FOREIGN KEY
	(
	CodCuenta
	) REFERENCES dbo.Cuentas
	(
	CodCuentas
	) ON UPDATE  CASCADE 
	 ON DELETE  CASCADE 
	
GO
ALTER TABLE dbo.Presupuesto WITH NOCHECK ADD CONSTRAINT
	FK_Presupuesto_Periodos FOREIGN KEY
	(
	NPeriodo
	) REFERENCES dbo.Periodos
	(
	NPeriodo
	) ON UPDATE  CASCADE 
	 ON DELETE  CASCADE 
	
GO
COMMIT
