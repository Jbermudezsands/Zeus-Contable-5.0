/*
   domingo, 30 de diciembre de 201207:43:11 a.m.
   Usuario: 
   Servidor: JUAN\SQL2005
   Base de datos: SistemaContableIpemsa
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
ALTER TABLE dbo.Conciliacion
	DROP CONSTRAINT DF_Conciliacion_Activo
GO
CREATE TABLE dbo.Tmp_Conciliacion
	(
	FechaConciliacion smalldatetime NOT NULL,
	CodCuenta nvarchar(50) NOT NULL,
	SaldoEstadoCuenta decimal(18, 2) NULL,
	Activo bit NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.Tmp_Conciliacion ADD CONSTRAINT
	DF_Conciliacion_Activo DEFAULT ((1)) FOR Activo
GO
IF EXISTS(SELECT * FROM dbo.Conciliacion)
	 EXEC('INSERT INTO dbo.Tmp_Conciliacion (FechaConciliacion, CodCuenta, SaldoEstadoCuenta, Activo)
		SELECT FechaConciliacion, CodCuenta, SaldoEstadoCuenta, Activo FROM dbo.Conciliacion WITH (HOLDLOCK TABLOCKX)')
GO
DROP TABLE dbo.Conciliacion
GO
EXECUTE sp_rename N'dbo.Tmp_Conciliacion', N'Conciliacion', 'OBJECT' 
GO
ALTER TABLE dbo.Conciliacion ADD CONSTRAINT
	PK_Conciliacion PRIMARY KEY CLUSTERED 
	(
	FechaConciliacion,
	CodCuenta
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
