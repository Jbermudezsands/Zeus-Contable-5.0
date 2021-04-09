/*
   miércoles, 7 de abril de 202111:11:50
   Usuario: sa
   Servidor: JUANBERMUDEZ
   Base de datos: SistemaContableRevesta
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
ALTER TABLE dbo.Reportes ADD
	Nperiodo float(53) NULL
GO
ALTER TABLE dbo.Reportes SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
