/*
   viernes, 20 de noviembre de 202003:42:24 p.m.
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
ALTER TABLE dbo.TransaccionesSolicitudPago ADD
	KeyPresupuesto nvarchar(250) NULL,
	Presupuesto nvarchar(250) NULL
GO
ALTER TABLE dbo.TransaccionesSolicitudPago SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
