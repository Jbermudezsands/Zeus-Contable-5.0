/*
   miércoles, 06 de enero de 202104:14:05 p.m.
   Usuario: sa
   Servidor: JUANBERMUDEZ\SQL2014
   Base de datos: SistemaContableEmtrides
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
ALTER TABLE dbo.IndiceSolicitudPago ADD
	NumeroTransaccion int NULL,
	NPeriodoTransaccion int NULL
GO
ALTER TABLE dbo.IndiceSolicitudPago SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
