/*
   jueves, 30 de abril de 202009:25:38 a.m.
   Usuario: 
   Servidor: JUANBERMUDEZ-PC\SQL2014
   Base de datos: SistemaContableMakuCafe
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
ALTER TABLE dbo.CordenadasCheque ADD
	Ciudad nvarchar(50) NULL
GO
ALTER TABLE dbo.CordenadasCheque ADD CONSTRAINT
	DF_CordenadasCheque_Ciudad DEFAULT N'Managua' FOR Ciudad
GO
ALTER TABLE dbo.CordenadasCheque SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
