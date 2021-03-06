/*
   miércoles, 19 de diciembre de 201812:10:50 p.m.
   Usuario: 
   Servidor: JUANBERMUDEZ\SQL2014
   Base de datos: Contabilidad1
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
ALTER TABLE dbo.ProyectosContratistas ADD
	Activo bit NULL
GO
ALTER TABLE dbo.ProyectosContratistas ADD CONSTRAINT
	DF_ProyectosContratistas_Activo DEFAULT 1 FOR Activo
GO
ALTER TABLE dbo.ProyectosContratistas SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
select Has_Perms_By_Name(N'dbo.ProyectosContratistas', 'Object', 'ALTER') as ALT_Per, Has_Perms_By_Name(N'dbo.ProyectosContratistas', 'Object', 'VIEW DEFINITION') as View_def_Per, Has_Perms_By_Name(N'dbo.ProyectosContratistas', 'Object', 'CONTROL') as Contr_Per 