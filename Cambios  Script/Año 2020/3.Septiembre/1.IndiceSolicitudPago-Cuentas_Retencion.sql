/*
   jueves, 08 de octubre de 202009:51:09 a.m.
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
ALTER TABLE dbo.IndiceSolicitudPago ADD
	CtaRetencion1 nvarchar(50) NULL,
	CtaRetencion2 nvarchar(50) NULL,
	CtaRetencion3 nvarchar(50) NULL,
	CtaRetencion4 nvarchar(50) NULL,
	CtaRetencion5 nvarchar(50) NULL,
	CtaRetencion6 nvarchar(50) NULL
GO
ALTER TABLE dbo.IndiceSolicitudPago SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
