/*
   miércoles, 06 de mayo de 202004:25:38 p.m.
   Usuario: 
   Servidor: JUANBERMUDEZ-PC\SQL2014
   Base de datos: SistemaContableCopam
   Aplicación: 
*/

/* Para evitar posibles problemas de pérdida de datos, debe revisar este script detalladamente antes de ejecutarlo fuera del contexto del diseñador de base de datos.*/
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
CREATE TABLE dbo.ContraCuentaPlanillaLeche
	(
	IdContraCuentaPlanilla int NOT NULL IDENTITY (1, 1),
	CuentaDebito nvarchar(50) NULL,
	CuentaCredito nvarchar(50) NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.ContraCuentaPlanillaLeche ADD CONSTRAINT
	PK_ContraCuentaPlanillaLeche PRIMARY KEY CLUSTERED 
	(
	IdContraCuentaPlanilla
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
ALTER TABLE dbo.ContraCuentaPlanillaLeche SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
