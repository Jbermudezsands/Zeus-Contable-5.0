/*
   miércoles, 05 de agosto de 202005:03:35 p.m.
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
ALTER TABLE dbo.NConsecutivos ADD
	ConstanciaRetencion float(53) NULL
GO
ALTER TABLE dbo.NConsecutivos ADD CONSTRAINT
	DF_NConsecutivos_ConsecutivoCheque DEFAULT 0 FOR ConsecutivoCheque
GO
ALTER TABLE dbo.NConsecutivos ADD CONSTRAINT
	DF_NConsecutivos_ConsecutivoImporta DEFAULT 0 FOR ConsecutivoImporta
GO
ALTER TABLE dbo.NConsecutivos ADD CONSTRAINT
	DF_NConsecutivos_ConstanciaRetencion DEFAULT 0 FOR ConstanciaRetencion
GO
ALTER TABLE dbo.NConsecutivos SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
