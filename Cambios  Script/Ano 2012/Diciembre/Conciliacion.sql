/*
   jueves, 06 de diciembre de 201206:35:30 a.m.
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
CREATE TABLE dbo.Conciliacion
	(
	FechaConciliacion smalldatetime NOT NULL,
	CodCuenta nvarchar(50) NULL,
	SaldoEstadoCuenta decimal(18, 2) NULL,
	Activo bit NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.Conciliacion ADD CONSTRAINT
	DF_Conciliacion_Activo DEFAULT 1 FOR Activo
GO
ALTER TABLE dbo.Conciliacion ADD CONSTRAINT
	PK_Conciliacion PRIMARY KEY CLUSTERED 
	(
	FechaConciliacion
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
