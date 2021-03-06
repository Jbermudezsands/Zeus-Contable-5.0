/*
   lunes, 22 de abril de 201903:50:13 p.m.
   Usuario: 
   Servidor: JUANBERMUDEZ\SQL2014
   Base de datos: SistemContablePanam
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
	Activo bit NULL,
	Procesado bit NULL
GO
ALTER TABLE dbo.IndiceSolicitudPago ADD CONSTRAINT
	DF_IndiceSolicitudPago_Activo DEFAULT 1 FOR Activo
GO
ALTER TABLE dbo.IndiceSolicitudPago ADD CONSTRAINT
	DF_IndiceSolicitudPago_Procesado DEFAULT 0 FOR Procesado
GO
ALTER TABLE dbo.IndiceSolicitudPago SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
