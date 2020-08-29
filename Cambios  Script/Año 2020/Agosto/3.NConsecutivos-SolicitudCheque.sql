/*
   lunes, 17 de agosto de 202004:35:42 p.m.
   Usuario: 
   Servidor: JUANBERMUDEZ-PC\SQL2014
   Base de datos: SistemaContableEMTRIDES
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
	ConsecutivoSolicitudCheque float(53) NULL
GO
ALTER TABLE dbo.NConsecutivos ADD CONSTRAINT
	DF_NConsecutivos_ConsecutivoSolicitudCheque DEFAULT 0 FOR ConsecutivoSolicitudCheque
GO
ALTER TABLE dbo.NConsecutivos SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
